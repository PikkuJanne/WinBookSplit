"""Finite authored transport controls; no book processing or shell evaluation."""

import json
import subprocess
import sys
import time


def emit(value, sequence=1):
    frame = {"protocol": "winbooksplit.interaction", "version": 1,
             "session": "authored-transport", "sequence": sequence,
             "stage": "input_ready", "mode": "manual", **value}
    sys.stdout.write("[WBS-INTERACTION] " + json.dumps(
        frame, ensure_ascii=False, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def terminal(message):
    print(json.dumps({"protocol": "winbooksplit.result", "version": 1,
        "mode": "manual", "status": "invalid_input", "code": "invalid_start_pages",
        "message": message, "warnings": [], "fallback_modes": [], "exit_code": 2,
        "written_count": 0, "execution": None}, ensure_ascii=False,
        separators=(",", ":")), flush=True)
    return 2


def main():
    kind = sys.argv[1]
    if kind == "descendant":
        time.sleep(8)
        return 0
    if kind == "handshake":
        print("Human before", flush=True)
        emit({"label": "Å日本 & % ! $(literal)"})
        first = json.loads(sys.stdin.readline())
        print("AUTHORED_SESSION_STDERR", file=sys.stderr, flush=True)
        emit({"stage": "plan_ready"}, 2)
        second = json.loads(sys.stdin.readline())
        print("\nHuman final no-newline")
        return terminal(json.dumps([first, second], ensure_ascii=False,
                                   separators=(",", ":")))
    if kind == "large-event":
        emit({"pad": "P" * 100000 + "Å日本"})
        answer = json.loads(sys.stdin.readline())
        print("Human after event", flush=True)
        return terminal(json.dumps(answer, ensure_ascii=False))
    if kind == "queue-flood":
        for sequence in range(1, 20):
            emit({}, sequence)
        time.sleep(8)
        return 0
    if kind == "oversized-event":
        emit({"pad": "P" * (8388608 + 1024)})
        time.sleep(8)
        return 0
    if kind in {"exact-event-cap", "one-over-event-cap"}:
        cap = 8388608 + (kind == "one-over-event-cap")
        frame = {"protocol": "winbooksplit.interaction", "version": 1,
                 "session": "authored-transport", "sequence": 1,
                 "stage": "input_ready", "mode": "manual", "pad": ""}
        empty = json.dumps(frame, separators=(",", ":"))
        frame["pad"] = "P" * (cap - len(empty.encode("utf-8")))
        payload = json.dumps(frame, separators=(",", ":"))
        assert len(payload.encode("utf-8")) == cap
        print("[WBS-INTERACTION] " + payload, flush=True)
        if kind == "one-over-event-cap":
            time.sleep(8)
            return 0
        json.loads(sys.stdin.readline())
        return terminal("exact cap reply received")
    if kind == "wait-with-descendant":
        child = subprocess.Popen([sys.executable, "-I", "-B", "-X", "utf8",
                                  __file__, "descendant"])
        emit({"owned_child_pid": child.pid})
        time.sleep(8)
        return 0
    if kind == "ignore-stdin":
        emit({})
        time.sleep(8)
        return 0
    if kind == "nul":
        return terminal("stdin EOF" if sys.stdin.readline() == "" else "unexpected input")
    raise ValueError("Unknown authored transport control")


if __name__ == "__main__":
    raise SystemExit(main())
