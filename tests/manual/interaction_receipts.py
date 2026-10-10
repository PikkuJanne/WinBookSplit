"""Strict recorded parent/engine handshakes; no console or viewer pass inferred."""

from hashlib import sha256
import json
import re


def need(value, message):
    if not value:
        raise ValueError(message)


def capture(text):
    def frames(prefix):
        return [json.loads(line[len(prefix):]) for line in text.splitlines() if line.startswith(prefix)]
    return {"engine_invocations": frames("[ENGINE] "),
            "requests": frames("[WBS-INTERACTION] "),
            "replies": frames("[INTERACTION-REPLY] "),
            "displayed_plans": frames("[PLAN] ")}


def validate(record, summaries, stages, actions, *, starts=None, execution=None):
    """Bind captured requests/replies to the one actual argv and terminal process."""
    need(isinstance(record, dict) and set(record) == {"engine_invocations", "requests", "replies", "displayed_plans"},
         "Exact interaction log capture absent")
    invocations, requests, replies = (record[name] for name in ("engine_invocations", "requests", "replies"))
    need(isinstance(invocations, list) and len(invocations) == 1 and isinstance(invocations[0], dict),
         "Exactly one actual engine invocation required")
    argv = invocations[0].get("arguments")
    need(isinstance(argv, list) and len(argv) >= 11 and argv[:4] == ["-I", "-B", "-X", "utf8"]
         and argv[9] == "--interactive" and isinstance(argv[10], str) and re.fullmatch(r"[0-9a-f]{32}", argv[10]),
         "Interactive argv/session nonce absent or in the wrong position")
    need(isinstance(requests, list) and isinstance(replies, list) and len(requests) == len(replies) == len(stages) == len(actions)
         and [event.get("stage") for event in requests] == stages and [reply.get("action") for reply in replies] == actions,
         "Exact interaction stages/actions/counts differ")
    need(isinstance(summaries, list) and len(summaries) == 1 and summaries[0].get("InteractionError") is None
         and summaries[0].get("InputError") is None and summaries[0].get("InputWriterStopped") is True
         and type(summaries[0].get("QueuedReplyCount")) is int and summaries[0]["QueuedReplyCount"] == len(replies)
         and type(summaries[0].get("InteractionCount")) is int
         and type(summaries[0].get("ReplyCount")) is int
         and summaries[0]["InteractionCount"] == len(requests)
         and summaries[0]["ReplyCount"] == len(replies), "Native interaction totals differ from recorded frames")
    current_mode = argv[7]
    plans = []
    for number, (event, reply) in enumerate(zip(requests, replies), 1):
        for frame in (event, reply):
            need(isinstance(frame, dict) and frame.get("protocol") == "winbooksplit.interaction"
                 and type(frame.get("version")) is int and frame["version"] == 1
                 and frame.get("session") == argv[10] and type(frame.get("sequence")) is int
                 and frame["sequence"] == number, "Interaction nonce/sequence/protocol differs")
        need(event.get("mode") == current_mode, "Interaction changed modes without an explicit retry")
        action = reply["action"]
        extra = {"starts"} if action == "starts" else {"mode"} if action == "retry" else {"plan_sha256"} if action == "execute" else set()
        need(set(reply) == {"protocol", "version", "session", "sequence", "action"} | extra,
             "Interactive reply has missing or unexpected data")
        if action == "starts":
            need(reply.get("starts") == starts and isinstance(starts, str), "Actual manual-start reply differs")
        if action == "retry":
            need(event["stage"] == "no_plan" and reply.get("mode") in event.get("fallback_modes", []), "Retry was not explicitly offered")
            current_mode = reply["mode"]
        if event["stage"] == "no_plan":
            result = event.get("result")
            need(isinstance(result, dict) and result.get("protocol") == "winbooksplit.result" and type(result.get("version")) is int and result["version"] == 1
                 and result.get("status") == "no_plan" and result.get("exit_code") == 5 and result.get("written_count") == 0
                 and result.get("execution") is None and result.get("mode") == event["mode"]
                 and event.get("fallback_modes") == result.get("fallback_modes"), "Nonterminal no-plan falsely claimed output")
        if event["stage"] == "plan_ready":
            plan, canonical = event.get("plan"), event.get("plan_json")
            need(isinstance(plan, dict) and isinstance(canonical, str) and json.loads(canonical) == plan
                 and canonical == json.dumps(plan, sort_keys=True, ensure_ascii=True, separators=(",", ":"))
                 and event.get("plan_sha256") == sha256(canonical.encode("utf-8")).hexdigest(), "Displayed immutable plan hash differs")
            plans.append(event)
            entries, coverage = plan.get("entries"), plan.get("coverage")
            need(plan.get("mode") == event["mode"] and type(plan.get("total_pages")) is int and plan["total_pages"] > 0
                 and isinstance(entries, list) and bool(entries) and isinstance(coverage, dict)
                 and coverage.get("complete") is True and type(coverage.get("covered_pages")) is int
                 and coverage["covered_pages"] == plan["total_pages"] and type(coverage.get("section_count")) is int
                 and coverage["section_count"] == len(entries)
                 and isinstance(plan.get("source_identity"), dict)
                 and isinstance(plan["source_identity"].get("sha256"), str)
                 and re.fullmatch(r"[0-9a-f]{64}", plan["source_identity"]["sha256"]), "Confirmed plan source/physical coverage absent")
            previous = 0
            for entry in entries:
                need(isinstance(entry, dict) and type(entry.get("start")) is int and type(entry.get("end")) is int
                     and entry["start"] == previous < entry["end"] <= plan["total_pages"], "Confirmed plan omitted or duplicated physical pages")
                previous = entry["end"]
            need(previous == plan["total_pages"], "Confirmed plan omitted final physical pages")
            need(action != "execute" or reply.get("plan_sha256") == event["plan_sha256"], "Consent did not bind the displayed plan")
    need(record.get("displayed_plans") == plans, "Recorded readable plan differs from the engine request")
    if execution is not None:
        need(len(plans) == 1 and actions[-1:] == ["execute"], "Published output lacks exact plan consent")
        plan = plans[0]["plan"]
        need(all(execution.get(field) == plan.get(field) for field in ("mode", "total_pages", "source_identity", "coverage"))
             and type(execution.get("written_count")) is int and execution["written_count"] == len(plan.get("entries", [])), "Execution changed confirmed source/coverage/count")
        entries, outputs = plan.get("entries"), execution.get("outputs")
        need(isinstance(entries, list) and isinstance(outputs, list) and len(entries) == len(outputs)
             and all(all(output.get(key) == value for key, value in entry.items()) for entry, output in zip(entries, outputs)),
             "Written sections differ from the confirmed plan")
    return argv


def require_noninteractive(record, summaries):
    need(isinstance(record, dict) and record.get("requests") == record.get("replies") == record.get("displayed_plans") == [],
         "Noninteractive control fabricated a plan or consent")
    invocations = record.get("engine_invocations")
    need(isinstance(invocations, list) and len(invocations) == 1 and "--interactive" not in invocations[0].get("arguments", []),
         "Noninteractive control used interactive argv")
    need(isinstance(summaries, list) and len(summaries) == 1
         and summaries[0].get("InteractionCount", 0) == summaries[0].get("ReplyCount", 0) == 0
         and summaries[0].get("InteractionError") is None, "Noninteractive control interaction totals differ")
