' clipper.vbs - clipboard watcher, VBScript port of index.hta
'
' Watches the clipboard for text matching a pattern and swaps it out for a
' replacement value. Same behaviour as the HTA, minus the browser: VBScript,
' WScript.Shell and PowerShell for the clipboard calls.
'
' Double-click to run. Stop it from Task Manager, or run with cscript:
'     cscript //nologo clipper.vbs

Option Explicit

' #START-CONFIG
Const PATTERN     = "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}\b"
Const REPLACEMENT = "your@value.corn"
Const INTERVAL_MS = 2500
' #END-CONFIG

Dim fso, shell, regex, tempPath
Set fso    = CreateObject("Scripting.FileSystemObject")
Set shell  = CreateObject("WScript.Shell")
Set regex  = New RegExp
tempPath   = fso.GetSpecialFolder(2) & "\clipper-" & fso.GetTempName() & ".txt"

regex.Pattern    = PATTERN
regex.IgnoreCase = True
regex.Global     = True

Main

Sub Main()
    ' "/once" does a single pass and exits: handy for a scheduled task, and
    ' the only way to test this without leaving a process running.
    ' Written as a nested If on purpose: VBScript's And does not
    ' short-circuit, so Arguments(0) would be read even with no arguments.
    Dim once, current, updated
    once = False
    If WScript.Arguments.Count > 0 Then
        If LCase(WScript.Arguments(0)) = "/once" Then once = True
    End If

    Do
        current = GetClipboard()
        If Len(current) > 0 Then
            updated = regex.Replace(current, REPLACEMENT)
            ' Only write back on a real change. Without this the loop would
            ' rewrite the clipboard every tick once the replacement matches
            ' the pattern itself.
            If StrComp(updated, current, vbBinaryCompare) <> 0 Then
                SetClipboard updated
            End If
        End If
        If once Then Exit Do
        WScript.Sleep INTERVAL_MS
    Loop
End Sub

Function GetClipboard()
    Dim proc, text
    GetClipboard = ""
    On Error Resume Next
    ' [Console]::Out.Write instead of a bare Get-Clipboard: stdout would
    ' otherwise gain a trailing newline, which the write-back would then
    ' add to the clipboard on every pass.
    Set proc = shell.Exec("powershell -NoProfile -Command ""[Console]::Out.Write((Get-Clipboard -Raw))""")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    text = proc.StdOut.ReadAll()
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    GetClipboard = text
End Function

Sub SetClipboard(text)
    ' Round-trip through a temp file so quotes and apostrophes in the text
    ' cannot break the PowerShell command line.
    Dim stream, cmd, exitCode
    On Error Resume Next

    Set stream = fso.CreateTextFile(tempPath, True, True)
    stream.Write text
    stream.Close

    cmd = "powershell -NoProfile -Command ""$t = Get-Content -Raw -LiteralPath '" & _
          tempPath & "'; Set-Clipboard -Value $t"""
    exitCode = shell.Run(cmd, 0, True)

    fso.DeleteFile tempPath, True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub
