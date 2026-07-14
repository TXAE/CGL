' ============================================================================
' SAP Login.vbs  – cleaned up & optimized
' Daniel Hermes – refactor April 2026
'TODO: Handle wrong user/password
'TODO: Handle scripting disabled by user
'
' Responsibilities:
'   - Detect existing SAP session
'   - Log in automatically if not logged in
'   - Return SAP session COM object for automarion of SAP GUI
'
' Notes:
'   - SAP GUI interaction stays 100% in VBScript
'   - PowerShell is used only for:
'       * Credential Manager
'       * Password dialog
'       * WM_SETTEXT credential injection
' ============================================================================

Option Explicit

' -------------------------
' Globals
' -------------------------
Dim shell, fso, session
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Dim target
target = "TERMSRV/ceberr55mp.eu.corp.cargill.com"


'WScript.Echo "WScript.ScriptFullName: " & WScript.ScriptFullName

' Only run SAPLogin()-function here if SAP Login.vbs was called standalone (not called from another script)
' Don't run SAPLogin()-function here if called from another script
' --- Other scripts:
'       - load SAP Login.vbs entirely
'       - set loadedFromMainScript = True (flag to indicate another script is calling SAP Login.vbs
'       - call SAPLogin()-function itself to get session-object (session object is used for automation of SAP GUI)
Dim loadedFromAnotherScript
If IsEmpty(loadedFromAnotherScript) Then Set session = SAPLogin()

' ============================================================================
' MAIN FUNCTION
' ============================================================================
Function SAPLogin()

    Dim session
    Set session = isLoggedIntoSAP()

    ' ---- Already logged in → return session immediately
    If Not session Is Nothing Then
        session.FindById("wnd[0]").Maximize
        Set SAPLogin = session
        Exit Function
    End If

    ' ---- Get credentials (PowerShell once)
    Dim username, password
    Call GetCredentials(username, password)

    'WScript.Echo Now() & "- Launching SAP login window..."
    Call LaunchSapFile(username)

    Call WaitForWindow(username, 10000)
    'WScript.Echo Now() & "- SAP login window active. Injecting credentials..."

    ' ---- Inject credentials (PowerShell once)
    Dim injectResult
    injectResult = InjectCredentials(username, password)
    If InStr(injectResult, "OK") = 0 Then
        WScript.Echo "SAP login injection returned: " & vbCrLf & injectResult
        WScript.Quit
    End If

    ' ---- Wait for Easy Access
    Call WaitForWindow("SAP Easy Access", 9000)

    ' ---- Attach to SAP session
    Set session = isLoggedIntoSAP()
    If session Is Nothing Then
        WScript.Echo "SAP login failed - no session after login"
        WScript.Quit
    End If

    session.FindById("wnd[0]").Maximize
    Set SAPLogin = session

End Function

' ============================================================================
' SAP SESSION DETECTION (unchanged, authoritative)
' ============================================================================
Function isLoggedIntoSAP()
    On Error Resume Next
    Set isLoggedIntoSAP = Nothing

    Dim SapGuiAuto, application, connection, session
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then Exit Function

    Set application = SapGuiAuto.GetScriptingEngine
    If application.Children.Count = 0 Then Exit Function

    Set connection = application.Children(0)
    If connection.Children.Count = 0 Then Exit Function

    Set session = connection.Children(0)
    Set isLoggedIntoSAP = session

    On Error GoTo 0
End Function

' ============================================================================
' CREDENTIAL HANDLING (PowerShell, single responsibility)
' ============================================================================
Sub GetCredentials(ByRef username, ByRef password)

    Dim output
    output = RunPowerShellFile(PS_ReadCred(target))
    'WScript.Echo "Credential Manager output: " & vbCrLf & output
    If InStr(output, "|") > 0 Then
        username = Split(output, "|")(0)
        password = Split(output, "|")(1)
        Exit Sub
    End If

    ' ---- No saved credential → ask for password
    username = CreateObject("WScript.Network").UserName
    password = RunPowerShell(PS_PromptPassword())

    If Len(password) = 0 Then
        WScript.Echo "No password entered."
        WScript.Quit
    End If

    ' ---- Save credential to credential manager for next time (PowerShell once)
    Dim writeCredResult
    writeCredResult = RunPowerShellFile(PS_WriteCred(target, username, password))
    If InStr(writeCredResult, "CredWrite failed") = 1 Then
        WScript.Echo writeCredResult
        WScript.Quit
    End If
End Sub

' ============================================================================
' SAP LOGIN STARTUP
' ============================================================================
Sub LaunchSapFile(username)
    Dim sapPath, f
    sapPath = shell.ExpandEnvironmentStrings("%TEMP%") & "\sap_autologin.sap"

    Set f = fso.CreateTextFile(sapPath, True)
    f.WriteLine "[System]"
    f.WriteLine "Name=PW1"
    f.WriteLine "Client=100"
    f.WriteLine "[User]"
    f.WriteLine "Name=" & username
    f.WriteLine "[Function]"
    f.WriteLine "Title=" & username
    f.Close

    shell.Run """" & sapPath & """", 1, False
End Sub

' ============================================================================
' WINDOW WAIT
' ============================================================================
Sub WaitForWindow(title, timeoutMs)
    Dim i, ok
    ok = False

    For i = 1 To (timeoutMs \ 100)
        ok = shell.AppActivate(title)
        If ok Then Exit Sub
        WScript.Sleep 100
    Next

    WScript.Echo Now() & " - Timeout waiting for window: " & title
    WScript.Quit
End Sub

' ============================================================================
' LOGIN FIELD INJECTION (PowerShell once)
' ============================================================================
Function InjectCredentials(title, password)
    InjectCredentials = RunPowerShellFile(PS_Inject(title, title, password))
End Function

' ============================================================================
' POWERSHELL INVOCATION (FAST – no temp file)
' ============================================================================
Function RunPowerShell(cmd)
    Dim exec, stdout, stderr
    
    Set exec = shell.Exec( _
        "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command " & Chr(34) & cmd & Chr(34))
    
    stdout = ""
    stderr = ""
    
    Do While Not exec.StdOut.AtEndOfStream
        stdout = stdout & exec.StdOut.ReadLine & vbCrLf
    Loop
    
    Do While Not exec.StdErr.AtEndOfStream
        stderr = stderr & exec.StdErr.ReadLine & vbCrLf
    Loop
    
    If Len(stderr) > 0 Then
        WScript.Echo "POWERSHELL STDERR:" & vbCrLf & stderr
    End If
    
    RunPowerShell = Trim(stdout)
End Function

' ============================================================================
' POWERSHELL INVOCATION (SLOW – with temp file) - TODO: Get rid of this if possible
' currently needed for CredRead & CrewdWrite which fails without temp file for some reason
' ============================================================================
Function RunPowerShellFile(psCode)
    Dim tempFolder, psPath, file, exec, output

    Set tempFolder = fso.GetSpecialFolder(2)
    psPath = tempFolder & "\sap_ps_" & Replace(Timer, ".", "") & ".ps1"

    Set file = fso.CreateTextFile(psPath, True)
    file.Write psCode
    file.Close

    Set exec = shell.Exec( _
        "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & psPath & """")

    output = exec.StdOut.ReadAll

    On Error Resume Next
    fso.DeleteFile psPath

    RunPowerShellFile = Trim(output)
End Function

' ============================================================================
' POWERSHELL CODE BUILDERS
' ============================================================================
Function PS_ReadCred(target)

    'works but lot of code - better version possible?
    PS_ReadCred = _
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
        "   $username = $cred.UserName -replace ""`r?`n"", """";" & vbCrLf & _
        "   Write-Output ($username + '|' + $pass);" & vbCrLf & _
        "   [CredMan]::CredFree($ptr)" & vbCrLf & _
        "} else {" & vbCrLf & _
        "   Write-Output 'Credential not found or access denied.'" & vbCrLf & _
        "}"

    ' old way to output username and password separately, switched to single line with delimiter for easier parsing in VBScript
    '"   Write-Output ('Username: ' + $cred.UserName);" & vbCrLf & _
    '"   Write-Output ('Password: ' + $pass);" & vbCrLf & _
End Function

Function PS_WriteCred(target, user, pw)
    'PowerShell-code saving credentials to credential manager
    PS_WriteCred = _
        "function Write-Credential {" & vbCrLf & _
        "    param (" & vbCrLf & _
        "        [string]$Target," & vbCrLf & _
        "        [string]$Username," & vbCrLf & _
        "        [string]$Password" & vbCrLf & _
        "    )" & vbCrLf & _
        "    Add-Type -TypeDefinition @'" & vbCrLf & _
        "using System;" & vbCrLf & _
        "using System.Runtime.InteropServices;" & vbCrLf & _
        "using System.Text;" & vbCrLf & _
        "public class CredMan {" & vbCrLf & _
        "    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]" & vbCrLf & _
        "    public struct CREDENTIAL {" & vbCrLf & _
        "        public int Flags;" & vbCrLf & _
        "        public int Type;" & vbCrLf & _
        "        public string TargetName;" & vbCrLf & _
        "        public string Comment;" & vbCrLf & _
        "        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;" & vbCrLf & _
        "        public int CredentialBlobSize;" & vbCrLf & _
        "        public IntPtr CredentialBlob;" & vbCrLf & _
        "        public int Persist;" & vbCrLf & _
        "        public int AttributeCount;" & vbCrLf & _
        "        public IntPtr Attributes;" & vbCrLf & _
        "        public string TargetAlias;" & vbCrLf & _
        "        public string UserName;" & vbCrLf & _
        "    }" & vbCrLf & _
        "    [DllImport(""advapi32.dll"", SetLastError = true, CharSet = CharSet.Unicode)]" & vbCrLf & _
        "    public static extern bool CredWrite([In] ref CREDENTIAL userCredential, [In] uint flags);" & vbCrLf & _
        "}" & vbCrLf & _
        "'@" & vbCrLf & _
        "    $cred = New-Object CredMan+CREDENTIAL" & vbCrLf & _
        "    $cred.Type = 1" & vbCrLf & _
        "    $cred.TargetName = $Target" & vbCrLf & _
        "    $cred.UserName = $Username" & vbCrLf & _
        "    $cred.Persist = 2" & vbCrLf & _
        "    $bytes = [System.Text.Encoding]::Unicode.GetBytes($Password)" & vbCrLf & _
        "    $cred.CredentialBlobSize = $bytes.Length" & vbCrLf & _
        "    $cred.CredentialBlob = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($bytes.Length)" & vbCrLf & _
        "    [System.Runtime.InteropServices.Marshal]::Copy($bytes, 0, $cred.CredentialBlob, $bytes.Length)" & vbCrLf & _
        "    $result = [CredMan]::CredWrite([ref]$cred, 0)" & vbCrLf & _
        "    [System.Runtime.InteropServices.Marshal]::FreeHGlobal($cred.CredentialBlob)" & vbCrLf & _
        "    if (-not $result) {" & vbCrLf & _
        "        Write-Output ('CredWrite failed with error code: ' + [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())" & vbCrLf & _
        "        exit 2" & vbCrLf & _
        "    } else {" & vbCrLf & _
        "        Write-Output 'Credential stored successfully.'" & vbCrLf & _
        "        exit 0" & vbCrLf & _
        "    }" & vbCrLf & _
        "}" & vbCrLf & _
        "Write-Credential -Target """ & target & """ -Username """ & user & """ -Password """ & pw & """"
End Function

Function PS_PromptPassword()
    PS_PromptPassword = _
        "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
        "$form = New-Object Windows.Forms.Form" & vbCrLf & _
        "$form.Text = 'Enter SAP password'" & vbCrLf & _
        "$form.Size = New-Object Drawing.Size(300,150)" & vbCrLf & _
        "$form.StartPosition = 'CenterScreen'" & vbCrLf & _
        "$form.KeyPreview = $true" & vbCrLf & _
        "$label = New-Object Windows.Forms.Label" & vbCrLf & _
        "$label.Text = 'Enter SAP password:'" & vbCrLf & _
        "$label.AutoSize = $true" & vbCrLf & _
        "$label.Location = New-Object Drawing.Point(10,20)" & vbCrLf & _
        "$form.Controls.Add($label)" & vbCrLf & _
        "$textbox = New-Object Windows.Forms.TextBox" & vbCrLf & _
        "$textbox.Location = New-Object Drawing.Point(10,50)" & vbCrLf & _
        "$textbox.Width = 260" & vbCrLf & _
        "$textbox.UseSystemPasswordChar = $true" & vbCrLf & _
        "$form.Controls.Add($textbox)" & vbCrLf & _
        "$okButton = New-Object Windows.Forms.Button" & vbCrLf & _
        "$okButton.Text = 'OK'" & vbCrLf & _
        "$okButton.Location = New-Object Drawing.Point(100,80)" & vbCrLf & _
        "$okButton.Add_Click({ $form.Tag = $textbox.Text; $form.Close() })" & vbCrLf & _
        "$form.Controls.Add($okButton)" & vbCrLf & _
        "$form.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $okButton.PerformClick() } })" & vbCrLf & _
        "$form.Tag = $null" & vbCrLf & _
        "$form.ShowDialog() | Out-Null" & vbCrLf & _
        "$pw = $form.Tag" & vbCrLf & _
        "If ([string]::IsNullOrWhiteSpace($pw)) { Write-Output 1 } else { Write-Output $pw }"
End Function

Function PS_Inject(title, u, p)
    PS_Inject = _
        "Add-Type -TypeDefinition @'" & vbCrLf & _
        "using System;" & vbCrLf & _
        "using System.Text;" & vbCrLf & _
        "using System.Runtime.InteropServices;" & vbCrLf & _
        "public static class W {" & vbCrLf & _
        " [DllImport(""user32.dll"", CharSet=CharSet.Auto, SetLastError=true)] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);" & vbCrLf & _
        " [DllImport(""user32.dll"", CharSet=CharSet.Auto, SetLastError=true)] public static extern bool IsWindowVisible(IntPtr hWnd);" & vbCrLf & _
        " [DllImport(""user32.dll"", CharSet=CharSet.Auto, SetLastError=true)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);" & vbCrLf & _
        " [DllImport(""user32.dll"", SetLastError=true)] public static extern int GetWindowTextLength(IntPtr hWnd);" & vbCrLf & _
        " [DllImport(""user32.dll"", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, string lParam);" & vbCrLf & _
        " [DllImport(""user32.dll"", SetLastError=true)] public static extern bool PostMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);" & vbCrLf & _
        " [DllImport(""user32.dll"", CharSet=CharSet.Auto, SetLastError=true)] public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);" & vbCrLf & _
        " public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);" & vbCrLf & _
        " public static IntPtr foundWindow = IntPtr.Zero;" & vbCrLf & _
        " public static string searchText = null;" & vbCrLf & _
        " public static bool EnumWindowCallback(IntPtr hwnd, IntPtr lParam) {" & vbCrLf & _
        "     if (!IsWindowVisible(hwnd)) return true;" & vbCrLf & _
        "     int len = GetWindowTextLength(hwnd);" & vbCrLf & _
        "     if (len == 0) return true;" & vbCrLf & _
        "     StringBuilder sb = new StringBuilder(len + 1);" & vbCrLf & _
        "     GetWindowText(hwnd, sb, sb.Capacity);" & vbCrLf & _
        "     if (sb.ToString().Contains(searchText)) { foundWindow = hwnd; return false; }" & vbCrLf & _
        "     return true;" & vbCrLf & _
        " }" & vbCrLf & _
        " public static IntPtr FindWindowByText(string text) {" & vbCrLf & _
        "     foundWindow = IntPtr.Zero;" & vbCrLf & _
        "     searchText = text;" & vbCrLf & _
        "     EnumWindows(EnumWindowCallback, IntPtr.Zero);" & vbCrLf & _
        "     return foundWindow;" & vbCrLf & _
        " }" & vbCrLf & _
        "}" & vbCrLf & _
        "'@;" & vbCrLf & _
        "$WM_SETTEXT = 0x000C; $WM_KEYDOWN = 0x0100; $WM_KEYUP = 0x0101; $VK_RETURN = 0x0D;" & vbCrLf & _
        "$title = '" & Replace(title, "'", "''") & "';" & vbCrLf & _
        "$user = '" & Replace(u, "'", "''") & "';" & vbCrLf & _
        "$pw = '" & Replace(p, "'", "''") & "';" & vbCrLf & _
        "Write-Output ('Searching top-level window by title: ' + $title);" & vbCrLf & _
        "$hWnd = [W]::FindWindowByText($title);" & vbCrLf & _
        "Write-Output ('FindWindowByText returned: ' + $hWnd.ToInt64());" & vbCrLf & _
        "if ($hWnd -eq [IntPtr]::Zero) { Write-Output 'NOTFOUND'; exit 1 }" & vbCrLf & _
        "$hEdit1 = [W]::FindWindowEx($hWnd, [IntPtr]::Zero, 'Edit', $null);" & vbCrLf & _
        "Write-Output ('FindWindowEx(Edit) returned: ' + $hEdit1.ToInt64());" & vbCrLf & _
        "if ($hEdit1 -eq [IntPtr]::Zero) { Write-Output 'NO_EDITS'; exit 2 }" & vbCrLf & _
        "$hEdit2 = [W]::FindWindowEx($hWnd, $hEdit1, 'Edit', $null);" & vbCrLf & _
        "Write-Output ('FindWindowEx(second Edit) returned: ' + $hEdit2.ToInt64());" & vbCrLf & _
        "if ($hEdit2 -eq [IntPtr]::Zero) { Write-Output 'USER_SET_ONLY'; exit 0 }" & vbCrLf & _
        "[W]::SendMessage($hEdit1, $WM_SETTEXT, [IntPtr]::Zero, $user) | Out-Null; [W]::SendMessage($hEdit2, $WM_SETTEXT, [IntPtr]::Zero, $pw) | Out-Null; [W]::PostMessage($hWnd, $WM_KEYDOWN, [IntPtr]$VK_RETURN, [IntPtr]0) | Out-Null; [W]::PostMessage($hWnd, $WM_KEYUP, [IntPtr]$VK_RETURN, [IntPtr]0) | Out-Null; Write-Output 'OK'; exit 0"
End Function