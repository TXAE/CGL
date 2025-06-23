' Daniel Hermes, 23. June 2025
' pulls ALL (including completed) repair orders (PM01) from yesterday to today (or the whole weekend if today is a Monday) in 1GOB plant & prints them
' For credentials to be used automatically, make sure that your SAP login is saved in Windows Credential Manager (Control Panel\User Accounts\Credential Manager)
' as a generic windows credential with the following target:
' This section pulls the pass from credential manager
target = "TERMSRV/ceberr55mp.eu.corp.cargill.com"

' Create PowerShell script content
psCode = _
"Add-Type -TypeDefinition @'" & vbCrLf & _
"using System;" & vbCrLf & _
"using System.Runtime.InteropServices;" & vbCrLf & _
"public class CredMan {" & vbCrLf & _
" [DllImport(""advapi32.dll"", SetLastError = true, CharSet = CharSet.Unicode)]" & vbCrLf & _
" public static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credentialPtr);" & vbCrLf & _
" [DllImport(""advapi32.dll"", SetLastError = true)]" & vbCrLf & _
" public static extern void CredFree(IntPtr buffer);" & vbCrLf & _
" [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]" & vbCrLf & _
" public struct CREDENTIAL {" & vbCrLf & _
"     public int Flags;" & vbCrLf & _
"     public int Type;" & vbCrLf & _
"     public string TargetName;" & vbCrLf & _
"     public string Comment;" & vbCrLf & _
"     public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;" & vbCrLf & _
"     public int CredentialBlobSize;" & vbCrLf & _
"     public IntPtr CredentialBlob;" & vbCrLf & _
"     public int Persist;" & vbCrLf & _
"     public int AttributeCount;" & vbCrLf & _
"     public IntPtr Attributes;" & vbCrLf & _
"     public string TargetAlias;" & vbCrLf & _
"     public string UserName;" & vbCrLf & _
" }" & vbCrLf & _
"}" & vbCrLf & _
"'@;" & vbCrLf & _
"$ptr = [IntPtr]::Zero;" & vbCrLf & _
"if ([CredMan]::CredRead('" & target & "', 1, 0, [ref]$ptr)) {" & vbCrLf & _
"   $cred = [System.Runtime.InteropServices.Marshal]::PtrToStructure($ptr, [Type][CredMan+CREDENTIAL]);" & vbCrLf & _
"   $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($cred.CredentialBlob, $cred.CredentialBlobSize / 2);" & vbCrLf & _
"   Write-Output ('Username: ' + $cred.UserName);" & vbCrLf & _
"   Write-Output ('Password: ' + $pass);" & vbCrLf & _
"   [CredMan]::CredFree($ptr)" & vbCrLf & _
"} else {" & vbCrLf & _
"   Write-Output 'Credential not found or access denied.'" & vbCrLf & _
"}"

rem "  Write-Output $pass;" & vbCrLf & _ rem Used earlier to return just the pass

Set shell = CreateObject("WScript.Shell")

' Write to temporary .ps1 file
Set fso = CreateObject("Scripting.FileSystemObject")
tempPath = shell.ExpandEnvironmentStrings("%TEMP%") & "\getcred.ps1"
Set file = fso.CreateTextFile(tempPath, True)
file.Write psCode
file.Close

' Run the PowerShell script

Set exec = shell.Exec("powershell -NoProfile -ExecutionPolicy Bypass -File """ & tempPath & """")
output = ""
Do While Not exec.StdOut.AtEndOfStream
   output = output & exec.StdOut.ReadLine() & vbCrLf
Loop
If InStr(output, "Credential not found or access denied.") = 1 Then
   MsgBox "Credential not found or access denied in credential manager for target: " & target
   WScript.Quit
End If

Dim username, password
username = ""
password = ""
lines = Split(output, vbCrLf)
For Each line In lines
   If InStr(line, "Username: ") = 1 Then
      username = Trim(Mid(line, Len("Username: ") + 1))
      rem cut domain (e.g. EU\) from username - SAP login does not use domain
      If InStr(username, "\") > 0 Then
         username = Split(username, "\")(1)
      End If
   ElseIf InStr(line, "Password: ") = 1 Then
      password = Trim(Mid(line, Len("Password: ") + 1))
   End If
Next
rem MsgBox "Username: " & username & vbCrLf & "Password: " & password

' delete the temp file
On Error Resume Next
fso.DeleteFile tempPath

'set path to temporary .sap file
sapFilePath = shell.ExpandEnvironmentStrings("%TEMP%") & "\temp_login.sap"

'write the .sap file contents
Set file = fso.CreateTextFile(sapFilePath, True)
file.WriteLine "[System]"
file.WriteLine "Name=PW1"
file.WriteLine "Description=PRD: PW1 ERP TC2"
file.WriteLine "Client=100"
file.WriteLine "[User]"
file.WriteLine "Name=" & username
file.WriteLine "[Function]"
file.WriteLine "Title=" & username
file.Close

'launch the .sap file
shell.Run """" & sapFilePath & """"

WaitForWindow(username)

shell.SendKeys "{TAB}"
shell.SendKeys password
shell.SendKeys "{ENTER}"

WaitForWindow("SAP Easy Access")

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   If Not application.Children.Count = 0 Then
      Set connection = application.Children(0)
   Else
      Set connection = application.OpenConnection("PW1", True)
      rem MsgBox "Logging in..."
      rem WScript.ConnectObject.Quit
   End If
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "IW38"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/chkDY_MAB").selected = true ' Order status Completed = TRUE
session.findById("wnd[0]/usr/chkDY_HIS").selected = true ' Order status Historical = TRUE
session.findById("wnd[0]/usr/ctxtAUART-LOW").text = "PM01" ' PM01 = repair orders (also known as work orders), PM02 are proactive orders (preventive maintenance)

rem this shows WOs that were changed in SAP during this time
session.findById("wnd[0]/usr/ctxtDATUB").text = "" ' Period Start
session.findById("wnd[0]/usr/ctxtDATUV").text = "" ' Period End

rem this shows WOs created during this time
Dim yesterday, start_date
yesterday = DateAdd("d", -1, Date)
rem WScript.Echo "Today: " & Date & vbCrLf & "Yesterday: " & yesterday
rem WScript.Echo "Weekday(yesterday): " & Weekday(yesterday)
rem MsgBox "Today: " & Date & vbCrLf &"Yesterday: " & yesterday
If Weekday(yesterday) = 1 Then
   rem WScript.Echo "Today is a Monday."
   start_date = DateAdd("d", -3, Date)
Else
   start_date = yesterday
   rem WScript.Echo "Today is not a Monday. Start_date: " & start_date
End If

session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = start_date
session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text = Date

session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = "1GOB"
rem session.findById("wnd[0]/usr/ctxtVARIANT").text = "TEST"
 

session.findById("wnd[0]/tbar[1]/btn[8]").press ' this executes the query


rem WScript.Sleep 7000 ' Wait for query to execute - not needed, but leaving here in case I need pause for demonstration purposes

session.findById("wnd[0]").sendVKey 86 rem this is CTRL+P to initiate print

rem this is necessary only if not set in System - User Profile - Own Data - Defaults - Print Immediately
rem session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM2").setFocus ' this select Properties - Print Time - Immediately
rem session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM2").key = "X"

rem this is necessary bc have not found out how to change defaults for these yet
session.findById("wnd[1]").sendVKey 6
rem format X_44_120 (ABAP/4 list: At least 44 rows by 120 columns) - makes it bigger than default Format - change if not everything fits
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "PAART","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "PAART","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "PAART","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-LINCT").text = "44"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/ctxtPRI_PARAMS-PAART").text = "X_44_120"
rem disable ALV Statisteks & Selections (do not need this extra page)
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").expandNode "spoolpostal"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "ALVST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "ALVST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").topNode = "PAART"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "ALVST","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/chkPRIPAR_DYN-ALVST").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/chkPRIPAR_DYN-ALVST").selected = false
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").selectItem "ALVSL","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").ensureVisibleHorizontalItem "ALVSL","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/cntlCUSTOM/shellcont/shell").doubleClickItem "ALVSL","Column2"
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/chkPRIPAR_DYN-ALVSL").setFocus
session.findById("wnd[2]/usr/tabsTABSTRIP/tabpTAB2/ssubSUBSCREEN:SAPLSPRI:0500/ssubSUBSCREEN:SAPLSPRI:0600/chkPRIPAR_DYN-ALVSL").selected = false
session.findById("wnd[2]/tbar[0]/btn[0]").press

rem this is to click "continue"-button and start the print
rem opens print dialog in background. TODO: find out how to open in foreground
session.findById("wnd[1]/tbar[0]/btn[13]").press 

' Wait a moment to make windows print dialog appear
rem WScript.Sleep 4000
rem WshShell.SendKeys "{ENTER}"





Sub WaitForWindow(windowTitle)
   Dim WshShell, windowFound, i, timeoutInMilliseconds
   Set WshShell = CreateObject("WScript.Shell")
   windowFound = False
   timeoutInMilliseconds = 8000

   For i = 1 To timeoutInMilliseconds
      If shell.AppActivate(windowTitle) Then
         rem WScript.Echo "Window with window title - " & windowTitle & " - found after roughly " & i & " ms."
         windowFound = True
         Exit For
      End If
      WScript.Sleep 1
   Next

   If not windowFound Then
      WScript.Echo "Window with window title - " & windowTitle & " - NOT found after roughly " & timeoutInMilliseconds & " ms."
      WScript.Quit
   End If
End Sub
