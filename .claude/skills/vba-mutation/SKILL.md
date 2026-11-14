---
name: vba-mutation
description: Prove that a new or changed VBA test (T_*, in vba/modTests.bas) really discriminates - break the code it covers on purpose, run vba-tests, see the test fail, restore the module. Use when AL asks for the mutation proof of a VBA test, or after adding a T_ test.
disable-model-invocation: true
---

# VBA mutation proof

CLAUDE.md: "prove that a new test discriminates: break the code on purpose (mutation) and watch it fail, then restore."
A test that stays green on broken code proves nothing. This skill does it for VBA; `mutate.py` keeps the good text and
refuses a mutation that does not apply exactly once.

## Steps (in order, stop at the first red step)

1. **State.** `tasklist | grep -c EXCEL` (note it: skill `excel-hygiene`), and
   `.venv/Scripts/python.exe .claude/skills/vba-mutation/mutate.py status` says no module is left mutated.
   Do not edit `src/` or `vba/` while the run goes (each run builds from `vba/*.bas`).
2. **Choose the mutation.** The smallest change that makes the code under test wrong: swap a branch, return a constant, drop
   a condition, change a comparison. One mutation per run. Name which tests must fail.
   Not a mutation: a change that does not compile (that proves nothing about the test), or one in `modTests.bas`.
3. **Apply.**
   `.venv/Scripts/python.exe .claude/skills/vba-mutation/mutate.py apply vba/modX.bas --old "<text once>" --new "<other>"`
   The old text is copied from the module as it is now (ASCII, one place; the script stops if it is found 0 or 2+ times).
4. **Run and restore in the same command**, so that an interrupted turn cannot leave the module broken:
   `(timeout 600 powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run.ps1 vba-tests > "$TEMP/vba_mut.log" 2>&1); .venv/Scripts/python.exe .claude/skills/vba-mutation/mutate.py restore vba/modX.bas; tail -n 12 "$TEMP/vba_mut.log"`
   For a quick run of a few tests: copy `modTests.bas`, cut `TestNames()`, run `vba-tests --offline`, restore with `cp`
   (CLAUDE.md, "A few VBA tests only"): never commit the cut list.
5. **Read the result.** Expected: `KO` lines naming the tests (`KO T_Name.n | area | label | ...`), and `missed markers`
   for the branch you broke. The proof is: the tests you named fail, others stay `OK`.
   - The named test is `OK`: it does not discriminate. Fix the test (a stronger `Check`), then redo from step 3.
   - `COMPILE_FAILED`: the mutation does not compile; choose another.
   - `RPC server unavailable` / the run lost Excel: the known flake, rerun (up to 3 times); it is no proof either way.
6. **Verify the restore.** `mutate.py status` says no module is left mutated, `git diff --stat vba/modX.bas` shows only the
   real change, `tasklist | grep -c EXCEL` is back to the start (stop a leftover by PID only, `excel-hygiene`).
7. **Report in French**, short sentences: the module, the mutation in one line, the tests that failed (or did not), the log
   path. Say plainly if a test stayed green.

## Never

- commit with a module mutated (`status` exits 1 when one is left; run it before every commit);
- mutate more than once per run (the failures could not be told apart);
- write the log to `out/` (use `$TEMP`).
