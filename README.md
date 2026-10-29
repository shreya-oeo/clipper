# Clipper

**A clipboard watcher for Windows.** Leave it running and it rewrites copied
text that matches a pattern. Built for swapping email addresses out of copied
content, but the pattern is yours to change.

Two versions, same behaviour:

| File | Runtime | Notes |
| --- | --- | --- |
| `index.hta` | HTA + JScript + WSH | Original 2023 version. Window-less. |
| `clipper.vbs` | VBScript + WSH | Console or double-click. No browser engine. |

## Usage

VBScript version:

```bat
cscript //nologo clipper.vbs        :: run and keep watching
cscript //nologo clipper.vbs /once  :: one pass, then exit
```

Or double-click `clipper.vbs` to run it windowed. Stop it from Task Manager.

HTA version: double-click `index.hta`.

## Configuration

Both files keep their settings in a marked block near the top. Edit and re-run;
there is nothing to compile.

```vbscript
' #START-CONFIG
Const PATTERN     = "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}\b"
Const REPLACEMENT = "your@value.corn"
Const INTERVAL_MS = 2500
' #END-CONFIG
```

- `PATTERN` is a VBScript `RegExp` pattern, matched case-insensitively.
- `REPLACEMENT` is what matched text becomes.
- `INTERVAL_MS` is how often the clipboard is checked.

## How it works

1. Poll the clipboard every 2.5 seconds via `WScript.Shell` and PowerShell.
2. Run the pattern over the text.
3. If anything matched, write the changed text back.

The clipboard is only written when the text actually changes, so the watcher
does not fight itself when the replacement also matches the pattern.

Text is handed to PowerShell through a temp file rather than interpolated into
the command line, so quotes and apostrophes in copied content cannot break it.

## Caveats

- Polling, not event-driven. There is a window of up to `INTERVAL_MS` where the
  clipboard still holds the original text.
- PowerShell startup makes each tick cost ~100 ms of CPU. Raise `INTERVAL_MS`
  if that matters.
- Windows only. Clipboard APIs here are WSH and PowerShell.

## License

MIT. See [LICENSE](LICENSE).
