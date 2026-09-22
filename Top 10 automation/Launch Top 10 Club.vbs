' Double-click this to open the Top 10 / Annexure K generator in your
' browser (pick which one to run with the toggle at the top of the page).
' It starts a local server (nothing leaves this PC) and opens the page.
' Closing the browser tab does NOT stop the server - that's fine, this
' script always closes any earlier copy of it before starting a fresh one
' (so double-clicking again after an update always runs the latest code),
' so you never need to touch Task Manager yourself.

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
serverPath = scriptDir & "\webapp\server.py"

' Stop any earlier copy of this exact server that might still be running
' in the background, so we never end up with two stale processes both
' listening on the same port (which caused inconsistent behaviour before).
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set procs = wmi.ExecQuery("SELECT CommandLine, ProcessId FROM Win32_Process WHERE Name = 'pythonw.exe'")
For Each p In procs
    If Not IsNull(p.CommandLine) Then
        If InStr(1, p.CommandLine, serverPath, 1) > 0 Then
            Set toKill = GetObject("winmgmts:\\.\root\cimv2:Win32_Process.Handle='" & p.ProcessId & "'")
            toKill.Terminate()
        End If
    End If
Next

Set shell = CreateObject("WScript.Shell")
shell.Run "pythonw """ & serverPath & """", 0, False
