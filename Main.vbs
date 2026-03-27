' Path to your SAP Logon executable
sapLogonPath = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
' Launch SAP Logon if it isn 't already running
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(sapLogonPath) Then
    shell.Run Chr(34) & sapLogonPath & Chr(34), 1, False
    ' Wait for SAP Logon to open
WScript.Sleep 2000 ' Adjust this delay if needed
Else
    MsgBox "SAP Logon executable not found at: " & sapLogonPath, vbCritical, "Error"
    WScript.Quit
End If
' Connect to SAP GUI Scripting Engine
On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <  > 0 Then
    MsgBox "SAP GUI Scripting not enabled or SAP Logon not running.", vbCritical, "Error"
WScript.Quit
End If
On Error GoTo 0
Set application = SapGuiAuto.GetScriptingEngine
    ' If application.Children.Count = 0 Then
        ' MsgBox "No SAP connections found. Please create a connection in SAP Logon.", vbCritical, "Error"
' WScript.Quit
' End If
' Open the first connection in the list (or modify to target a specific connection)
Set connection = application.OpenConnection("SYSTEM_NAME", True)
' Wait for the SAP GUI session to initialize
WScript.Sleep 2000 ' Adjust this delay if needed
' Perform login steps
Set session = connection.Children(0)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "CLIENT_ID" ' Client
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "USERNAME" ' Username
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "PASSWORD" ' Password
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = Len("**")
    session.findById("wnd[0]").sendVKey 0