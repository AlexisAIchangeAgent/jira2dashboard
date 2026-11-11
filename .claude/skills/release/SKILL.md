---
name: release
description: Publish a release of jira2dashboard (zip of the code + built workbook, GitHub release) with every rule of CLAUDE.md, stopping to ask AL before the outward-facing -Publish.
disable-model-invocation: true
---

# Release

A release is outward-facing: it publishes the code and the workbook on GitHub. AL chooses the number and says yes to the
publication itself. Talk to AL in French, short simple sentences.

## Checklist (in order, stop at the first red step)

1. **Number.** Ask AL for `vX.Y.Z` (only this form; `v0.0.0-dryrun` is refused). The published ones are listed in
   CLAUDE.md (section Commands, "are published"). The tags are on GitHub: `git fetch --tags` before reading them.
2. **State.** `git status --porcelain` empty, `git rev-parse HEAD` == `git rev-parse origin/main` (push first),
   `gh auth status` logged in. No `EXCEL.EXE` of a run left (`tasklist | grep -c EXCEL`).
3. **Full pass** (skill `/full-pass`): coverage, vba-tests, test-excel, test-network, build. A known Excel crash is
   rerun, not ignored.
4. **Dry run:** `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/release.ps1 -Version v0.0.0`
   (log `scripts\logs\release.log`). Check in `out\release\v0.0.0`: the zip leaves out the `export-ignore` paths
   (`CLAUDE.md`, `archive/`, `.claude/`, `.venv/`), `notes.md` gives the MD5 and SHA256, and the workbook holds the
   version (hidden name `J2D_Version`: `unzip -p <xlsm> xl/workbook.xml | grep J2D_Version` shows `v0.0.0 commit ...`).
   Then delete `out\release\v0.0.0`.
5. **Ask AL** with the number, the commit and what the dry run showed. Publish only on a clear yes for this number.
6. **Publish:** `scripts/release.ps1 -Version vX.Y.Z -Publish`. It builds the workbook with `--version vX.Y.Z`, so the RUN
   line of its logs reads `jira2dashboard vX.Y.Z commit ...`.
7. **After:** add the version, the date and the commit to the line "are published" of CLAUDE.md, commit and push. Give AL
   the URL of the release and the MD5 of `jira2dashboard.xlsm` (for the allow-list of the endpoint protection: the file
   name and folder last longer than the hash).

## Never

- publish without AL's yes for this exact number, or reuse a number already tagged;
- write the vendor or product name of the endpoint protection in the notes (say EPDR or "endpoint protection");
- put a client value (BaseUrl, Email, ApiToken, proxy, Filter, iterations) in the workbook or the zip: they stay empty.
