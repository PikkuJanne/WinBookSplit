"""Transparent invented handshakes for validator units, never native evidence."""

from copy import deepcopy
from hashlib import sha256
import json


def make(execution, invocation, stages, actions, *, starts="2,3", no_plan=None):
    nonce = invocation["arguments"][10]
    requests, replies, plans = [], [], []
    mode = invocation["arguments"][7]
    for number, (stage, action) in enumerate(zip(stages, actions), 1):
        common = {"protocol": "winbooksplit.interaction", "version": 1, "session": nonce, "sequence": number}
        event, reply = {**common, "mode": mode, "stage": stage}, {**common, "action": action}
        if stage == "no_plan":
            event.update(result=deepcopy(no_plan), fallback_modes=deepcopy(no_plan["fallback_modes"]))
        if stage == "plan_ready":
            plan = {key: deepcopy(execution.get(key)) for key in ("mode", "total_pages", "source_identity", "coverage")}
            plan["entries"] = [{key: deepcopy(value) for key, value in output.items() if key not in {"sha256", "size_bytes", "page_count"}}
                               for output in execution["outputs"]]
            canonical = json.dumps(plan, sort_keys=True, ensure_ascii=True, separators=(",", ":"))
            event.update(plan=plan, plan_json=canonical, plan_sha256=sha256(canonical.encode()).hexdigest())
            plans.append(deepcopy(event))
        if action == "starts":
            reply["starts"] = starts
        if action == "retry":
            reply["mode"] = mode = "manual"
        if action == "execute":
            reply["plan_sha256"] = event["plan_sha256"]
        requests.append(event)
        replies.append(reply)
    return {"engine_invocations": [deepcopy(invocation)], "requests": requests, "replies": replies, "displayed_plans": plans}
