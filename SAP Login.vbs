' Daniel Hermes, 07. July 2025
' After this script has finished, SAP session should be there.
' Either user is already logged in, then easy.
' Otherwise check for saved credentials and try to log in with them.
' If no credentials are saved, script asks user to save credentials

'TODO: Handle wrong user/password
'TODO: Handle scripting disabled by user

Option Explicit ' forces to declare all variables with Dim, Private, or Public
Dim loadedFromAnotherScript, target, shell, fso, SapGuiAuto, application, connection, session, psCode, output, username, password, sapFilePath, file, dump

If IsEmpty(loadedFromAnotherScript) Then
    ' --- Only run this if not called from another script --
    'WScript.Echo "Called directly: " & WScript.ScriptFullName
    Set session = SAPLogin()
    'WScript.Echo "SAP session ID: " & session.Id
Else
    ' --- Don't run SAPLogin() here if called from another script
    ' The other script loads this script entirely & will call SAPLogin()-function itself to get session-object
    'WScript.Echo "Loaded from another script: " & WScript.ScriptFullName
End If

Function SAPLogin()
    target = "TERMSRV/ceberr55mp.eu.corp.cargill.com"
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set session = isLoggedIntoSAP()
    If session Is Nothing Then
        ' PowerShell-code checking if credential is already saved in credential manager
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
        output = RunPowerShellScript(psCode)
        username = ""
        password = ""
        If InStr(output, "Username: ") = 1 Then
            ' credential is already saved in credential manager - parse username and password from output
            Dim lines, line
            lines = Split(output, vbCrLf)
            For Each line In lines
                If InStr(line, "Username: ") = 1 Then
                    username = Trim(Mid(line, Len("Username: ") + 1))
                    ' cut domain(e.g.EU \ ) from username - SAP login does Not use domain
                    If InStr(username, "\") > 0 Then
                        username = Split(username, "\")(1)
                    End If
                ElseIf InStr(line, "Password: ") = 1 Then
                    password = Trim(Mid(line, Len("Password: ") + 1))
                End If
            Next
            'MsgBox "Username: " & username & vbCrLf & "Password: " & password
        Else
            'MsgBox   "Credential not found or access denied in credential manager for target: " & target & vbCrLf & _
            '         "Opening CredentialManager so you can check and maybe enter. "
            'shell.Run "control /name Microsoft.CredentialManager"
            'WScript.Sleep 1000 ' Wait for the window to open
            Dim objNetwork
            Set objNetwork = CreateObject("WScript.Network")
            username = objNetwork.UserName
            
            ' PowerShell-code asking user to input pw
            psCode = _
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
            
            password = RunPowerShellScript(psCode)
            If InStr(password, "1") = 1 Then
                WScript.Echo "user did not enter a pw - terminating script"
                WScript.Quit
            End If
            
            'PowerShell-code saving credentials to credential manager
            psCode = _
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
                "Write-Credential -Target """ & target & """ -Username """ & username & """ -Password """ & password & """"
            
            Dim credStoreResult
            credStoreResult = RunPowerShellScript(psCode)
            If InStr(credStoreResult, "CredWrite failed") = 1 Then
                WScript.Echo credStoreResult
                WScript.Quit
            End If
        End If
        
        'set path to temporary .sap file
        sapFilePath = shell.ExpandEnvironmentStrings("%TEMP%") & "\temp_login.sap"
        
        'write the .sap file contents
        Set file = fso.CreateTextFile(sapFilePath, True)
        file.WriteLine "[System]"
        file.WriteLine "Name=PW1" 'QW3"
        'file.WriteLine "Description=PRD: PW1 ERP TC2"
        file.WriteLine "Client=100" 'needed, otherwise will always get script-warning-popup from SAP!
        file.WriteLine "[User]"
        file.WriteLine "Name=" & username
        file.WriteLine "[Function]"
        file.WriteLine "Title=" & username
        file.Close
        
        'launch the .sap file
        shell.Run """" & sapFilePath & """"
        dump = Now & "- Launched SAP login window with .sap file for user: " & username & " - waiting for login window..."
        
        'wait for the login window to appear
        WaitForWindow(username)
        dump = dump & vbCrLf & Now & "- SAP login window appeared."
        
        ' Try to set username/password programmatically (WM_SETTEXT) and post Enter.
        Dim setResult
        setResult = TrySetLoginFields(username, username, password)

        If InStr(setResult, "OK") > 0 Then
            ' fields set and Enter posted by PowerShell
            dump = dump & vbCrLf & Now & "- TrySetLoginFields result: " & setResult
            dump = dump & vbCrLf & Now & "- Successfully set username and password into SAP login window and posted Enter via PowerShell. Now waiting for SAP Easy Access window..."
        Else
            dump = dump & vbCrLf & Now & "- ERROR! TrySetLoginFields result: " & setResult
            WScript.Echo dump
            WScript.Quit

            ' fallback to previous SendKeys approach - not working reliably, because SAP login window often does not get focused properly, so keys are sent to the wrong window
            ' ALSO RISKY BECAUSE IF SOMEONE HAS ANOTHER WINDOW WITH FOCUS, PASSWORD WOULD BE SENT THERE! DO NOT USE!
            'shell.SendKeys "{TAB}"
            'shell.SendKeys password
            'shell.SendKeys "{ENTER}"
        End If

        WaitForWindow("SAP Easy Access")

        Set session = isLoggedIntoSAP()
        If session Is Nothing Then
            WScript.Echo "Auto login did not work :("
            WScript.Quit
        End If
    End If
    
    session.findById("wnd[0]").maximize
    Set SAPLogin = session

    dump = dump & vbCrLf & Now & "- Done!"
    'WScript.Echo dump ' ENABLE THIS TO DEBUG
End Function

'Checks if user is logged into SAP. Returns the session-object if is logged in, NULL if not logged in
Function isLoggedIntoSAP()
    'Check if SAP is already running
    Set isLoggedIntoSAP = Nothing
    On Error Resume Next
    ' Try to get SAP GUI
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number = 0 Then
        
        ' Try to get the scripting engine
        Set application = SapGuiAuto.GetScriptingEngine
        If Err.Number = 0 Then
            
            ' Check if any connections exist
            If application.Children.Count > 0 Then
                
                ' Get the first connection
                Set connection = application.Children(0)
                If Err.Number = 0 Then
                    
                    ' Check if any sessions exist
                    If connection.Children.Count > 0 Then
                        
                        ' Get the first session
                        Set session = connection.Children(0)
                        If Err.Number = 0 Then
                            
                            ' Check if session is active
                            'If session.Info.IsLowSpeedConnection = False Then
                            Set isLoggedIntoSAP = session
                            'WScript.Echo "User is already logged in to SAP."
                            'Else
                            '    WScript.Echo "SAP session found, but not fully connected."
                            'End If
                            
                        End If
                    Else
                        'WScript.Echo "SAP login window is open, but user is not logged in."
                    End If
                End If
            Else
                'WScript.Echo "SAP GUI is open, but no connections found (login window not open)."
            End If
        Else
            'WScript.Echo "Failed to get SAP scripting engine."
        End If
    Else
        'WScript.Echo "SAP GUI is not running."
    End If
    On Error GoTo 0
End Function

Sub WaitForWindow(WindowTitle)
    Dim startTime, elapsedTime, timeoutInMilliseconds
    Dim activeTitle, psCode
    timeoutInMilliseconds = 9000
    startTime = Timer
    
    Do While True
        ' Attempt to bring any window whose MainWindowTitle contains WindowTitle to the foreground,
        ' then return the current foreground window title so we can verify focus.
        psCode = _
            "Add-Type -TypeDefinition @'" & vbCrLf & _
            "using System; using System.Runtime.InteropServices;" & vbCrLf & _
            "public class U {" & vbCrLf & _
            " [DllImport(""user32.dll"")] public static extern bool SetForegroundWindow(IntPtr hWnd);" & vbCrLf & _
            " [DllImport(""user32.dll"")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);" & vbCrLf & _
            "}" & vbCrLf & _
            "'@;" & vbCrLf & _
            "$target = '" & Replace(WindowTitle, "'", "''") & "';" & vbCrLf & _
            "$proc = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -and $_.MainWindowTitle -like (""*"" + $target + ""*"") } | Select -First 1;" & vbCrLf & _
            "if ($proc) { $h = $proc.MainWindowHandle; [U]::ShowWindow($h,5) | Out-Null; [U]::SetForegroundWindow($h) | Out-Null }" & vbCrLf & _
            "Add-Type -TypeDefinition @'" & vbCrLf & _
            "using System; using System.Runtime.InteropServices; using System.Text;" & vbCrLf & _
            "public class Win {" & vbCrLf & _
            " [DllImport(""user32.dll"")] public static extern IntPtr GetForegroundWindow();" & vbCrLf & _
            " [DllImport(""user32.dll"", CharSet=CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);" & vbCrLf & _
            "}" & vbCrLf & _
            "'@;" & vbCrLf & _
            "$sb = New-Object System.Text.StringBuilder 1024;" & vbCrLf & _
            "$hwnd = [Win]::GetForegroundWindow();" & vbCrLf & _
            "if ($hwnd -eq [IntPtr]::Zero) { Write-Output '' } else { [Win]::GetWindowText($hwnd,$sb,$sb.Capacity) | Out-Null; Write-Output $sb.ToString() }"

        activeTitle = RunPowerShellScript(psCode)

        ' Compare case-insensitive and allow partial match (titles often include extra text)
        If Len(activeTitle) > 0 Then
            If InStr(LCase(activeTitle), LCase(WindowTitle)) > 0 Then
                Exit Sub
            End If
        End If

        ' Check if timeout exceeded (convert Timer output to milliseconds)
        elapsedTime = (Timer - startTime) * 1000
        If elapsedTime >= timeoutInMilliseconds Then
            Exit Do
        End If

        WScript.Sleep 100
    Loop

    WScript.Echo "App with title - " & WindowTitle & " - NOT found in foreground after roughly " & timeoutInMilliseconds & " ms."
    WScript.Quit
End Sub

Function RunPowerShellScript(psCode)
    Dim tempFolder, psFile, psPath, exec, output, line

    ' Temporäre Datei erstellen
    Set tempFolder = fso.GetSpecialFolder(2) ' 2 = TemporaryFolder
    psPath = tempFolder & "\temp_script_" & Timer & ".ps1"

    Set psFile = fso.CreateTextFile(psPath, True)
    psFile.Write psCode
    psFile.Close

    ' PowerShell-Skript ausführen und Ausgabe lesen
    Set exec = shell.Exec("powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & psPath & """")

    output = ""
    Do While Not exec.StdOut.AtEndOfStream
        line = exec.StdOut.ReadLine
        output = output & line & vbCrLf
    Loop

    ' Temporäre Datei löschen
    On Error Resume Next
    fso.DeleteFile psPath

    RunPowerShellScript = Trim(output)
End Function

' Attempts to set username/password into the login window edit controls using WM_SETTEXT
' Returns the raw output from the PowerShell script (e.g. 'OK', 'USER_SET_ONLY', 'NO_EDITS', 'NOTFOUND')
Function TrySetLoginFields(WindowTitle, u, p)
    Dim psCode
    psCode = _
        "Add-Type -TypeDefinition @'" & vbCrLf & _
        "using System;" & vbCrLf & _
        "using System.Text;" & vbCrLf & _
        "using System.Runtime.InteropServices;" & vbCrLf & _
        "public static class W {" & vbCrLf & _
        " [DllImport(""user32.dll"", CharSet=CharSet.Auto, SetLastError=true)] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);" & vbCrLf & _
        " [DllImport(""user32.dll"", CharSet=CharSet.Auto, SetLastError=true)] public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);" & vbCrLf & _
        " [DllImport(""user32.dll"", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, string lParam);" & vbCrLf & _
        " [DllImport(""user32.dll"", SetLastError=true)] public static extern bool PostMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);" & vbCrLf & _
        "}" & vbCrLf & _
        "'@;" & vbCrLf & _
        "$WM_SETTEXT = 0x000C; $WM_KEYDOWN = 0x0100; $WM_KEYUP = 0x0101; $VK_RETURN = 0x0D;" & vbCrLf & _
        "$title = '" & Replace(WindowTitle, "'", "''") & "';" & vbCrLf & _
        "$user = '" & Replace(u, "'", "''") & "';" & vbCrLf & _
        "$pw = '" & Replace(p, "'", "''") & "';" & vbCrLf & _
        "$hWnd = [W]::FindWindow($null, $title);" & vbCrLf & _
        "if ($hWnd -eq [IntPtr]::Zero) { $proc = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -and $_.MainWindowTitle -like ('*' + $title + '*') } | Select -First 1; if ($proc) { $hWnd = $proc.MainWindowHandle } }" & vbCrLf & _
        "if ($hWnd -eq [IntPtr]::Zero) { Write-Output 'NOTFOUND'; exit 1 }" & vbCrLf & _
        "$hEdit1 = [W]::FindWindowEx($hWnd, [IntPtr]::Zero, 'Edit', $null);" & vbCrLf & _
        "if ($hEdit1 -ne [IntPtr]::Zero) { [W]::SendMessage($hEdit1, $WM_SETTEXT, [IntPtr]::Zero, $user) | Out-Null; $hEdit2 = [W]::FindWindowEx($hWnd, $hEdit1, 'Edit', $null); if ($hEdit2 -ne [IntPtr]::Zero) { [W]::SendMessage($hEdit2, $WM_SETTEXT, [IntPtr]::Zero, $pw) | Out-Null; [W]::PostMessage($hWnd, $WM_KEYDOWN, [IntPtr]$VK_RETURN, [IntPtr]0) | Out-Null; [W]::PostMessage($hWnd, $WM_KEYUP, [IntPtr]$VK_RETURN, [IntPtr]0) | Out-Null; Write-Output 'OK'; exit 0 } else { Write-Output 'USER_SET_ONLY'; exit 0 } } else { Write-Output 'NO_EDITS'; exit 2 }"

    TrySetLoginFields = RunPowerShellScript(psCode)
End Function