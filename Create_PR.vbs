rem Daniel Hermes, 29. April 2025
rem This script needs you to be in the "SAP Easy Access" menu. 
rem That is the menu just after login. You need to be logged in, but have no T-Code open.
rem Script will create a PR
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
session.findById("wnd[0]/tbar[0]/okcd").text = "ME51N"
session.findById("wnd[0]").sendVKey 0 ' Enter key

rem ENTER INFORMATION HERE
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text = _
"header note" + vbCr + "new line?" ' "header note" + vbCr + "new line?"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"EPSTP",_
"" ' leave empty for standard (physical thing), put D for service
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"TXZ01",_
"Short text" ' Put Text description
rem session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "TXZ01"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"MENGE",_
"69" ' Put quantity
rem session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter

rem for all GL Accounts:  https://cargillonline.sharepoint.com/:x:/r/sites/SAPSUBERLIN/Shared%20Documents/Allgemeines/Cost%20Centers%20-%20GL-Accounts%20-%20Vendor%20List/Chart%20of%20Account%20Mapping%20CCC%26Ghent%26Vilvoorde.xlsx?d=w2965b20c5ea943ad8636dde0fe3407a8&csf=1&web=1&e=90LgCe
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").text = _
"66210010" ' GL Account in SAP for REPAIRS & MAINTENANCE EXPENSE: 66210010

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = _
"109210797"
rem for all Cost centers: https://cargillonline.sharepoint.com/:x:/r/sites/SAPSUBERLIN/Shared%20Documents/Allgemeines/Cost%20Centers%20-%20GL-Accounts%20-%20Vendor%20List/NOVA%20W2%20Cost%20center%20structure_JDE%20vs%20SAP%201.xlsx?d=wb413017034b4402bb7ebcd93cbc58b56&csf=1&web=1&e=RdQG32
rem Person responsible: Ulrike Doerre U952336
rem 109210803 - 1GOB A995 LIQ CHOC LN 1 MILK/LITTLE DARK
rem 109210804 - 1GOB A995 LIQUID CHOCOLATE LINE 2 DARK
rem 109210805 - 1GOB A995 LIQUID CHOCOLATE LINE 3 WHITE
rem 109210806 - 1GOB A995 SOLID CHOCOLATE LINE 1
rem 109210807 - 1GOB A995 SOLID CHOCOLATE LINE 2

rem Person responsible: ASsaduzzaman Noor A467903
rem 109210797 - 1GOB A995 MAINTENANCE
rem 109210814 - 1GOB A995 CCC PLANT SYSTEMS-CONTROLS
rem 109210966 - 1GOB A995 BERLIN R WAREHOUSE PLANT 2

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell 0,"MEINS",_
"EA" ' EA = Each
