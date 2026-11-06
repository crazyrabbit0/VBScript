# VBScript

Small VBScript launchers for Windows. Each takes a file as its first argument, so they are meant for shortcuts, the **Send To** menu or other tools that pass a path.

## Scripts

| Script | What it does |
|---|---|
| `Run App  (Hidden).vbs` | Runs the given program or script with its window hidden. |
| `Run App  (Wait + Hidden).vbs` | Same, but waits for it to finish and returns its exit code. |
| `Run App  (Admin + Hidden).vbs` | Runs it hidden and elevated (shows the UAC prompt). |
| `Open Shortcut Location.vbs` | Opens Explorer with the file selected; for a `.lnk` shortcut it selects the shortcut's target instead. |

## Usage

```bat
wscript.exe "Run App  (Hidden).vbs" "C:\path\to\program.bat"
```

`.Script Shortcut.lnk` is a template shortcut that runs `wscript.exe "Script.vbs"`: copy it, point it at the script you want, and pass the target as an extra argument. For **Send To**, put a shortcut to the `.vbs` in `shell:sendto`.

## Issues

Found a bug or have an idea? [Open an issue](https://github.com/crazyrabbit0/VBScript/issues/new/choose).

## License

[GPL-3.0](LICENSE) © CrazyRabbit
