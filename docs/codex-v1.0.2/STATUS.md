# v1.0.2 status

V102-02 is complete; public v1.0.2 is not yet published.

Shared policy 3 accepts reviewed regular 3.15.0 and numeric later stable patches; all older floors and dependency locks retained. Run the actual 3.15 application/setup/process suite using the separate existing developer fixture driver.

Current task: V102-03. Checkpoint branch: `codex/v1.0.2-setup-python315`. Source/evidence references are in TASKS.json. Each commit is followed by normal push and fresh clean/live equality verification outside its own evidence commit.

Required native Windows 10/11 final-ZIP Python 3.15 gates remain NOT_RUN until actual execution. No historical result substitutes for these gates.
