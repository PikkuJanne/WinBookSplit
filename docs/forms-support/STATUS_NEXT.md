# Form support status

Checkpoint 1 is complete on `codex/forms-v1.0.5`. The
[checkpoint evidence](CHECKPOINT_1.md) records the released v1.0.4 failure and
the executable Windows regression. The application and VERSION remain at v1.0.4.

Next is checkpoint 2 of [CODEX_TASK.md](CODEX_TASK.md). Classify forms in the
shared PDF reader, pass harmless metadata through preview and split, and give
nonempty XFA a distinct reason. Keep widget-bearing PDFs behind the fail-closed
boundary until checkpoint 3 can preserve their appearances. The latest owner
PDF was unavailable, so its form type remains unknown.

No owner action is needed. Continue with this prompt:

> Continue PikkuJanne/WinBookSplit v1.0.5 form support in the actual checkout at
> D:\projects\WinBookSplit-main on codex/forms-v1.0.5. Read AGENTS.md and
> docs/forms-support/STATUS_NEXT.md and its linked evidence. Complete checkpoint
> 2 only. Implement bounded shared classification, make the harmless metadata
> cases pass native preview and split regressions, and distinguish nonempty XFA
> and malformed structures without publishing any widget-bearing output that
> has not been preserved. Keep the v1.0.4 setup, discovery, input and output
> protections. Run available automated tests, record results, commit, push,
> verify the live branch matches clean local HEAD, then hand over and stop.
