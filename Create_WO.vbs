' Daniel Hermes, 26. August 2025
Option Explicit ' forces to declare all variables with Dim, Private, or Public
Dim loadedFromMainScript, fso, file, code, session
loadedFromMainScript = True ' flag to indicate this script is calling the login script
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("SAP Login.vbs", 1)
code = file.ReadAll
file.Close
ExecuteGlobal code
Set session = SAPLogin()

session.StartTransaction "IW31"
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtAUFPAR-PM_AUFART").text = "PM01" 'repair order
'session.findById("wnd[0]/usr/cmbCAUFVD-PRIOK").key = "3" 'Priority
session.findById("wnd[0]/usr/subOBJECT:SAPLCOIH:7100/ctxtCAUFVD-TPLNR").text = "1GOB" 'Berlin Reinickendorf PPM Plus plant code
session.findById("wnd[0]/usr/subOBJECT:SAPLCOIH:7100/ctxtCAUFVD-EQUNR").text = "1001289969" 'software service schokolinie
session.findById("wnd[0]").sendVKey 0 'press enter

'have to do this again in case user was not logged in when launching the script. Not necessary when user already logged in, but keeping for stability
session.findById("wnd[0]/usr/cmbCAUFVD-PRIOK").key = "3" 'Priority
session.findById("wnd[0]").sendVKey 0 'press enter

Dim userInput
userInput = InputBox("Worum gehts bei der WO?", "Worum gehts bei der WO?")
If IsEmpty(userInput) Then
    MsgBox("no description entered. Terminating script...")
    WScript.Quit
End If
Dim objNetwork
Set objNetwork = CreateObject("WScript.Network")
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell").text = _
    userInput + vbCr + "WO geschrieben von " + objNetwork.UserName
session.findById("wnd[0]").sendVKey 11 ' Save
MsgBox(session.findById("wnd[0]/sbar").text)

session.StartTransaction "IW32"
session.findById("wnd[0]").sendVKey 0 'press enter
session.findById("wnd[0]").maximize
' set Functional Location back to 1GOB and delete Equipment to make it quicker to select a different functional location & equipment than set by default 
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subOBJECT:SAPLCOIH:7100/ctxtCAUFVD-TPLNR").text = "1GOB"
session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subOBJECT:SAPLCOIH:7100/ctxtCAUFVD-EQUNR").text = ""
session.findById("wnd[0]").sendVKey 0 'press enter