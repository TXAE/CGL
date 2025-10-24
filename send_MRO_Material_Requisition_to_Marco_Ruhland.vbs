' Daniel Hermes, 23. October 2025
' pull automated purchase requisitions of last 7 days from SAP and send via email to specific recipients
Option Explicit ' forces to declare all variables with Dim, Private, or Public
Dim loadedFromMainScript, fso, file, code, session
loadedFromMainScript = True ' flag to indicate this script is calling the login script
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("SAP Login.vbs", 1)
code = file.ReadAll
file.Close
ExecuteGlobal code
Set session = SAPLogin()

session.StartTransaction "ME5A" ' Purchase Requisitions: List Display


session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "1GOB"
session.findById("wnd[0]/usr/chkP_ERLBA").selected = true ' Include Closed Requisitions
session.findById("wnd[0]/usr/txtP_AFNAM").text = "MRO Material"
session.findById("wnd[0]").sendVKey 16 ' Dynamic Selections (Shift+F4) to filter for "Requisition Date"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").expandNode "          1" ' expand the first node in dynamic selections, called "Purchase Requisition"
' necessary to scroll down to node 28 called "Requisition Date"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "         28"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "         20"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "         28" ' double click node called "Requisition Date" 
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN001_%_APP_%-VALU_PUSH").press ' press button "Multiple Selection" to be able to filter for a date range
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select ' Select tab "Select Ranges"
Dim one_week_ago : one_week_ago = DateAdd("d", - 7, Date)
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,0]").text = one_week_ago ' Lower Limit date
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").text = Date ' Upper Limit date
session.findById("wnd[0]").sendVKey 8 ' Copy (F8) the dynamic selection
session.findById("wnd[0]").sendVKey 8 ' Execute (F8)


session.findById("wnd[0]").sendVKey 31 ' Mail Recipient (CTRL+F7)
session.findById("wnd[0]/usr/subSENDSCREEN:SAPLSO04:1020/subOBJECT:SAPLSO33:2300/tabsSO33_TAB1/tabpTAB1/ssubSUB1:SAPLSO33:2100/cntlEDITOR/shellcont/shell").text = _
    "Hallo Marco," + vbCr + "" + vbCr + "Nataliya wollte, dass Du jeden Freitag um 10 Uhr diese Liste mit den automatischen (vom SAP-System angefragten) Bestellungen der letzten Woche bekommst." + vbCr + _
    "Diese Liste wird automatisch generiert und Dir auch automatisch mithilfe eines Skriptes versandt." + vbCr + "" + vbCr + "Geniess Dein Wochenende!" + vbCr + "Daniel Hermes"

session.findById("wnd[0]/usr/subSENDSCREEN:SAPLSO04:1020/subRECLIST:SAPLSO04:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:SAPLSO04:0150/tblSAPLSO04REC_CONTROL/ctxtSOS04-L_ADR_NAME[0,0]").text = "M245834" ' Marco Ruhland
session.findById("wnd[0]/usr/subSENDSCREEN:SAPLSO04:1020/subRECLIST:SAPLSO04:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:SAPLSO04:0150/tblSAPLSO04REC_CONTROL/ctxtSOS04-L_ADR_NAME[0,1]").text = "n736374" ' Nataliya Hristova
session.findById("wnd[0]/usr/subSENDSCREEN:SAPLSO04:1020/subRECLIST:SAPLSO04:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:SAPLSO04:0150/tblSAPLSO04REC_CONTROL/ctxtSOS04-L_ADR_NAME[0,2]").text = "f699011" ' Fabian Erdmann
session.findById("wnd[0]/usr/subSENDSCREEN:SAPLSO04:1020/subRECLIST:SAPLSO04:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:SAPLSO04:0150/tblSAPLSO04REC_CONTROL/ctxtSOS04-L_ADR_NAME[0,3]").text = "g650208" ' George-Pierre Bulus
session.findById("wnd[0]").sendVKey 0 'enter
session.findById("wnd[0]").sendVKey 20 ' Send... (Shift+F8)