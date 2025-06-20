rem Daniel Hermes, 20. June 2025
rem file:///C:/Users/d363973/Cargill%20Inc/SAP%20BERLIN%20-%20Documents/Allgemeines/Anleitungen/Anwendung%20-%20Indirect%20Procurement/Anleitung_Anwendung_SAP_Inbound%20Delivery%20&%20Goods%20Movement.pdf
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   If Not application.Children.Count = 0 Then
      Set connection = application.Children(0)
   Else ' SAP probably not open
      MsgBox "SAP probably not logged in"
   End If
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
rem this stuff should normally always be the same
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "VL31N"
session.findById("wnd[0]").sendVKey 0 ' Enter key

rem ENTER DATA HERE
PurchaseOrder = "33456178"
ExternalID = "OUBZB-025535"
MeansOfTransport = "LKW B-AT-9744"











ItemText = "entsprechend Angebot " & Angebot & vbCr & _
PreisNetto & " EUR netto" & vbCr & _
"requisitioned by Daniel Hermes, D363973, Cargill Berlin-Reinickendorf chocolate factory" & vbCr & _
"+49 162 3478656"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text = _
ItemText
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"EPSTP",_
Service
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"TXZ01",_
"Angebot " & Angebot ' Put Text description
rem session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "TXZ01"
rem session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"MENGE",_
"1" ' Put quantity

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"MEINS",_
"EA" ' EA = Each

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"PREIS",_
PreisNetto ' net (netto) valuation Price

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"EEIND",_
DateAdd("d", 14, Date) ' Delivery Date - today in X days by default - enter custom date as DD.MM.YYYY

rem session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"WGBEZ",_
rem "010100" ' Material Group - Material Group Description - Description 2 for material group
rem               010095      -        01 PARTS             - MRO: PARTS
rem               010100      -        01 TOOLS             - MRO: TOOLS

rem SHOULD ALWAYS BE THE SAME
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"EKORG",_
"P003" ' Purchasing Organization - P003 = EMEA (Europe Middle East Asia)
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"NAME1",_
"BERLIN 11 DE PPM REINICK 1092" ' Plant
session.findById("wnd[0]").sendVKey 0 ' Enter key - switches to Item number 1


rem for all GL Accounts:  https://cargillonline.sharepoint.com/:x:/r/sites/SAPSUBERLIN/Shared%20Documents/Allgemeines/Cost%20Centers%20-%20GL-Accounts%20-%20Vendor%20List/Chart%20of%20Account%20Mapping%20CCC%26Ghent%26Vilvoorde.xlsx?d=w2965b20c5ea943ad8636dde0fe3407a8&csf=1&web=1&e=90LgCe
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").text = _
"66210010" ' GL Account in SAP for REPAIRS & MAINTENANCE EXPENSE: 66210010

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = _
"109210797"
rem for all Cost centers: https://cargillonline.sharepoint.com/:x:/r/sites/SAPSUBERLIN/Shared%20Documents/Allgemeines/Cost%20Centers%20-%20GL-Accounts%20-%20Vendor%20List/NOVA%20W2%20Cost%20center%20structure_JDE%20vs%20SAP%201.xlsx?d=wb413017034b4402bb7ebcd93cbc58b56&csf=1&web=1&e=RdQG32
rem Person responsible: Ulrike Doerre U952336
rem 109210803 - 1GOB A995 LIQ CHOC LN 1 MILK/LITTLE DARK
rem 109210804 - 1GOB A995 LIQUID CHOCOLATE LINE 2 DARK
rem 109210805 - 1GOB A995 LIQUID CHOCOLATE LINE 3 WHITE
rem 109210806 - 1GOB A995 SOLID CHOCOLATE LINE 1
rem 109210807 - 1GOB A995 SOLID CHOCOLATE LINE 2

rem Person responsible: Assaduzzaman Noor A467903
rem 109210797 - 1GOB A995 MAINTENANCE
rem 109210814 - 1GOB A995 CCC PLANT SYSTEMS-CONTROLS
rem 109210966 - 1GOB A995 BERLIN R WAREHOUSE PLANT 2

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13").select
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell").text = _
ItemText

rem EXPERIMENTAL - selects the following:
rem         Material Group - Material Group Description - Description 2 for material group
rem               010100      -        01 TOOLS             - MRO: TOOLS
rem CSSC Code Tools - but dont know how to change that afterwards ?! :(
rem 
rem session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"WGBEZ","010100"
rem session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
rem session.findById("wnd[1]/usr/lbl[1,11]").setFocus
rem session.findById("wnd[1]/usr/lbl[1,11]").caretPosition = 6
rem session.findById("wnd[1]").sendVKey 2