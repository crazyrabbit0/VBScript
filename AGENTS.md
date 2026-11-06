# AGENTS.md

Small Windows VBScript launchers, public on GitHub (`crazyrabbit0/VBScript`, branch `main`). No build, no tests; `README.md` lists every script.

## Conventions

- Each `.vbs` reads its target from `WScript.Arguments(0)` and quotes it with `"""" & path & """"` so paths with spaces work.
- File names use two spaces before the parenthesised variant, e.g. `Run App  (Hidden).vbs`; keep that pattern for new variants.
- `.Script Shortcut.lnk` is a binary template (`wscript.exe "Script.vbs"`); `.gitattributes` keeps `.lnk` files binary and text files LF.

## Checking a change

There is no syntax checker in CI. Test by hand on Windows: `cscript //nologo "<script>.vbs" "<a test file>"`, and check the window style, elevation prompt or Explorer selection the script promises.

## Commits

Conventional Commits, e.g. `feat(run-app): ...`, `docs: ...`. Update the README table when a script is added, removed or changes purpose.

Contributor-facing rules live in `CONTRIBUTING.md` and `.github/pull_request_template.md`; keep them in step with this file.
