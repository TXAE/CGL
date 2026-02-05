' Daniel Hermes, started 1. July 2025
' Mechanics & electricians have to log their work in a spreadsheet (called "Schichtbuch" in German) & in SAP. It is double work.
' This script parses through the spreadsheet, detects logged work in the spreadsheet that is not yet logged in SAP
' and can log the work in SAP (either fully automatic or asks user to confirm every single SAP interaction, depending on parameter)
Option Explicit ' forces to declare all variables with Dim, Private, or Public
Dim g_logFilePath, logFile, filePath, excelApp, workbook, sheet1, sheet2, lastRow, lastCol, loadedFromMainScript, fso, file, code, session, autoConfirmResponse, argFilePath, argUseCurrentExcel, argAutoConfirm
Dim sheet1_cached, sheet2_cached, g_timezoneBias, g_statusBuffer(), prevScreenUpdating, prevCalculation, column_in_excel_where_to_put_message, done_text_from_excel, cancelled_text_from_excel, SAP_plantcode
initialize()

' DEBUGGING SETTINGS
rowsToInspect = Array() 'PUT HERE WHICH ROWS YOU WANT TO CHECK. example: rowsToInspect = Array(6, 9, 69)
onlyParse_rowsToInspect = False 'TRUE IF YOU ONLY WANT THE SCRIPT TO PARSE THROUGH rowsToInspect. FALSE IF YOU WANT THE SCRIPT TO PARSE THROUGH ALL ROWS BUT LOG SPECIFIC STUFF FOR rowsToInspect




' === LOOP THROUGH ROWS ===
Dim i, j, SkipReason, rowText, Tag, WO_Nr, Schicht, Standort, Mitarbeiter, Bemerkung, Fehlerbeschreibung, Massnahme, Startzeit, Endzeit, DauerInH, Status, emptyRowCounter, rowsToInspect, onlyParse_rowsToInspect
For i = 2 To lastRow ' Assuming row 1 is header
    ' Ensure status buffer exists to collect per-row messages (buffered write to Excel)
    If Not IsArray(g_statusBuffer) Then
        ReDim g_statusBuffer(lastRow)
        Dim idxInit
        For idxInit = 0 To lastRow
            g_statusBuffer(idxInit) = ""
        Next
    End If
    If Not onlyParse_rowsToInspect Or IsInArray(i, rowsToInspect) Then
        SkipReason = ""
        rowText = ""
        For j = 1 To lastCol
            rowText = rowText & GetCellFromCache(sheet1_cached, 1, j) & ": " & GetCellFromCache(sheet1_cached, i, j) & vbCrLf
        Next 'col
        Tag = GetCellFromCache(sheet1_cached, i, 1)
        WO_Nr = GetCellFromCache(sheet1_cached, i, 3)
        Mitarbeiter = GetCellFromCache(sheet1_cached, i, 5)
        Bemerkung = GetCellFromCache(sheet1_cached, i, 7)
        Fehlerbeschreibung = GetCellFromCache(sheet1_cached, i, 8)
        Massnahme = GetCellFromCache(sheet1_cached, i, 9)
        Startzeit = GetCellFromCache(sheet1_cached, i, 10)
        Endzeit = GetCellFromCache(sheet1_cached, i, 11)
        DauerInH = GetCellFromCache(sheet1_cached, i, 12)
        Status = GetCellFromCache(sheet1_cached, i, 15)

        ' HARD skip conditions - no WO, already something in message column from a previous script execution or no status/cancelled
        If WO_Nr = "" Then
            SkipReason = SkipReason & vbCrLf & " - No WO found."
        ElseIf Len(WO_Nr) <> 9 Then
            SkipReason = SkipReason & vbCrLf & " - WO " & WO_Nr & " is " & Len(WO_Nr) & " characters long (but should be 9 characters long)."
        ElseIf Left(WO_Nr, 1) <> "4" Then
            SkipReason = SkipReason & vbCrLf & " - WO " & WO_Nr & " does not start with a '4.' "
        End If
        If GetCellFromCache(sheet1_cached, i, 16) <> "" Then
            SkipReason = SkipReason & vbCrLf & " -" & GetCellFromCache(sheet1_cached, i, 16)
        End If
        If Status = "" Then
            SkipReason = SkipReason & vbCrLf & " - No 'Status' for WO " & WO_Nr & " found."
        ElseIf Status = cancelled_text_from_excel And SkipReason = "" Then
            Log "Row " & i & " - WO " & WO_Nr & " is marked as cancelled in shift logbook. Proceeding to cancel in SAP..."
            ' Cancel the WO in SAP
            SafeStartTransaction "IW32"
            SafeSetText "wnd[0]/usr/ctxtCAUFVD-AUFNR", WO_Nr
            SafeSendVKey "wnd[0]", 0
            ' Set User Status
            SafePress "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btn%#AUTOTEXT001"
            ' Select CNCL Cancelled
            SafeSetSelected "wnd[1]/usr/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,2]", True
            SafeSendVKey "wnd[1]", 0 ' Enter
            SafeSendVKey "wnd[0]", 11 ' CTRL+S saves the order
            Check_if_WO_needs_TECO(wo_Nr)
            Dim cancelled_msg : cancelled_msg = "WO cancelled in SAP."
            SkipReason = SkipReason & vbCrLf & " - " & cancelled_msg
            Log cancelled_msg
            WriteToExcel i, column_in_excel_where_to_put_message, cancelled_msg, False
            ' Performance improvement: - problem: when script faults, not all is written! :(
            ' Buffer the write (will flush at the end). Also keep log output.
            'g_statusBuffer(i) = cancelled_msg
        End If
        
        'SOFT skip conditions - missing data that is needed for confirmation
        If Mitarbeiter = "" Then
            SkipReason = SkipReason & vbCrLf & " - No employee for WO " & WO_Nr & " found."
        End If
        If Tag = "" Then
            SkipReason = SkipReason & vbCrLf & " - No day for WO " & WO_Nr & " found."
        End If
        Dim StartzeitExcelFractionConverterToTime, StartzeitSafeCDate, StartzeitConvertToUTC, EndzeitExcelFractionConverterToTime, EndzeitSafeCDate, EndzeitConvertToUTC
        If Startzeit = "" Then
            SkipReason = SkipReason & vbCrLf & " - No Startzeit for WO " & WO_Nr & " found."
        Else
            StartzeitExcelFractionConverterToTime = ConvertExcelFractionToTime(Startzeit) 'before something like "0,3125" from Excel, after something like "07:30:00"
            StartzeitSafeCDate = SafeCDate(Tag & " " & StartzeitExcelFractionConverterToTime) 'before a Time like "07:30:00", after a Date like "01.07.2025 07:30:00"
            StartzeitConvertToUTC = ConvertToUTC(StartzeitSafeCDate) 'before Date in local time, after Date in UTC
        End If
        If Endzeit = "" Then
            SkipReason = SkipReason & vbCrLf & " - No Endzeit for WO " & WO_Nr & " found."
        Else
            EndzeitExcelFractionConverterToTime = ConvertExcelFractionToTime(Endzeit)
            EndzeitSafeCDate = SafeCDate(Tag & " " & EndzeitExcelFractionConverterToTime)
            EndzeitConvertToUTC = ConvertToUTC(EndzeitSafeCDate)
            If Startzeit <> "" And StartzeitConvertToUTC > EndzeitConvertToUTC Then
                Log "StartzeitConvertToUTC: " & StartzeitConvertToUTC & " is later than EndzeitConvertToUTC: " & EndzeitConvertToUTC & vbCrLf & _
                    "Making Endzeit one day later..."
                'example: user entered a single day, but starting time 10PM & finish time 3AM.
                EndzeitConvertToUTC = EndzeitConvertToUTC + 1 'make Endzeit one day later
            End If
        End If
        
        If IsInArray(i, rowsToInspect) Then
            Log vbCrLf & _
                "______________________________________________________________________________" & vbCrLf & _
                "Reached row " & i & " to inspect..." & vbCrLf & _
                "Startzeit initially (excel fraction) : " & Startzeit & vbCrLf & _
                "StartzeitExcelFractionConverterToTime: " & StartzeitExcelFractionConverterToTime & vbCrLf & _
                "StartzeitSafeCDate                   : " & StartzeitSafeCDate & vbCrLf & _
                "StartzeitConvertToUTC                : " & StartzeitConvertToUTC & vbCrLf & _
                "Endzeit initially (excel fraction) : " & Endzeit & vbCrLf & _
                "EndzeitExcelFractionConverterToTime: " & EndzeitExcelFractionConverterToTime & vbCrLf & _
                "EndzeitSafeCDate                   : " & EndzeitSafeCDate & vbCrLf & _
                "EndzeitConvertToUTC                : " & EndzeitConvertToUTC & vbCrLf
            'CleanupAndTerminate "Terminating due to reaching a row to inspect (debugging)..."
        End If
        

        If DauerInH = "" Then
            SkipReason = SkipReason & vbCrLf & " - No 'DauerInH' for WO " & WO_Nr & " found."
        Else
            DauerInH = ConvertExcelFractionToTime(DauerInH)
            If IsNull(DauerInH) Then
                SkipReason = SkipReason & vbCrLf & " - Invalid or missing 'DauerInH' for WO " & WO_Nr & "."
            End If
        End If
        If IsCachedRowEmpty(i) Then

            Dim emptyRowsUntilDone : emptyRowsUntilDone = 10
            emptyRowCounter = emptyRowCounter + 1
            Log "Row " & i & " detected as empty row. emptyRowCounter: " & emptyRowCounter
            If emptyRowCounter = emptyRowsUntilDone Then
                Log emptyRowsUntilDone & " empty rows detected."
                If argUseCurrentExcel = "yes" And argAutoConfirm = "yes" Then
                    Log "argUseCurrentExcel = yes And argAutoConfirm = yes, so copying upcoming PMs to shift logbook..."

                    If SAP_plantcode = "" Then
                        SafeStartTransaction "IW33"
                        ' can use any WO to get to the plant code, so don't put any WO here - this will SAP use the last one
                        SafeSendVKey "wnd[0]", 0
                        SAP_plantcode = SafeGetText("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/txtCAUFVD-IWERK")
                    End If
                    
                    SafeStartTransaction "IW38"
                    Dim today, tomorrow, Basic_Start_Date_LOW, Basic_Start_Date_HIGH
                    today = Date
                    tomorrow = DateAdd("d", 1, today)
                    Basic_Start_Date_LOW = tomorrow
                    If Weekday(today) = 6 Then ' today is Friday
                        Basic_Start_Date_HIGH = DateAdd("d", 3, today)
                    Else
                        Basic_Start_Date_HIGH = tomorrow
                    End If

                    SafeSetText "wnd[0]/usr/ctxtGSTRP-LOW", Basic_Start_Date_LOW ' Basic Start Date - LOW
                    SafeSetText "wnd[0]/usr/ctxtGSTRP-HIGH", Basic_Start_Date_HIGH ' Basic Start Date - HIGH
                    SafeSetText "wnd[0]/usr/ctxtDATUV", "" ' Period - LOW
                    SafeSetText "wnd[0]/usr/ctxtDATUB", "" ' Period - HIGH
                    SafeSetText "wnd[0]/usr/ctxtSWERK-LOW", SAP_plantcode
                    SafeSendVKey "wnd[0]", 8 ' Execute (F8)
                    
                    Dim grid, rowCount, colCount, ord, r, c
                    Set grid = SafeFindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
                    rowCount = grid.RowCount
                    'colCount = grid.ColumnCount

                    ' --- 1) Get the ColumnOrder as a collection of technical IDs ---
                    On Error Resume Next
                    Set ord = grid.ColumnOrder ' to get techId of column header (e.g. AUFNR = Order)
                    If Err.Number <> 0 Then
                        Err.Clear
                        CleanupAndTerminate "This ALV does not expose ColumnOrder as a collection. Cannot proceed without tech IDs."
                    End If
                    On Error GoTo 0

                    ' --- 2) Build a map: techId -> visible index ---
                    Dim techToIdx : Set techToIdx = CreateObject("Scripting.Dictionary")
                    Dim colId
                    For c = 0 To ord.Count - 1
                        colId = ord.Item(c)
                        If Not techToIdx.Exists(colId) Then techToIdx.Add colId, c
                    Next

                    ' --- 3) Define the required fields and their Excel target columns ---
                    Dim wanted : Set wanted = CreateObject("Scripting.Dictionary")
                    '   tech id     -> Excel column index (example mapping; adjust as you need)
                    wanted.Add "GSTRP", 1 ' Basic Start Date
                    wanted.Add "AUFNR", 3 ' Order
                    wanted.Add "ZZTIDNR", 4 ' Technical IdentNo.
                    wanted.Add "KTEXT", 7 ' Description

                    ' --- 4) Validate the layout: all required tech IDs must exist ---
                    Dim missing : Set missing = CreateObject("Scripting.Dictionary")
                    Dim k
                    For Each k In wanted.Keys
                        If Not techToIdx.Exists(k) Then missing.Add k, True
                    Next

                    If missing.Count > 0 Then
                        Dim msg, key
                        msg = "Required column(s) missing in current IW38 layout: "
                        For Each key In missing.Keys
                            msg = msg & key & " "
                        Next
                        msg = msg & vbCrLf & "Please load the correct ALV layout (e.g., with technical names) and try again."
                        CleanupAndTerminate msg
                    End If

                    ' --- 5) Pre-locate the KTEXT column index (for skip-row check) ---
                    Dim ktextIdx
                    ktextIdx = techToIdx("KTEXT")

                    ' --- 6) Main loop: Skip rows where KTEXT contains "Indutec", else process wanted fields only ---
                    Dim row_factor_to_adust_for_skip : row_factor_to_adust_for_skip = 1
                    For r = 0 To rowCount - 1
                        ' Read KTEXT strictly by tech ID; aborts if not readable
                        value = GetCellValueStrict(grid, r, "KTEXT")
                        
                        If InStr(value, "Indutec") = 0 Then
                            Dim row_in_excel_where_to_put : row_in_excel_where_to_put = i - emptyRowsUntilDone + row_factor_to_adust_for_skip + r
                            ' Process only the fields we want
                            For Each k In wanted.Keys
                                Dim value, column_in_excel_where_to_put : column_in_excel_where_to_put = wanted(k)

                                ' Read strictly by tech ID; aborts if not readable
                                value = GetCellValueStrict(grid, r, k)

                                If value <> "" Then
                                    Log "  Taking from SAP" & vbCrLf & _
                                        "  row: " & r & ", techId: " & k & ", value: " & value & vbCrLf & _
                                        "  and putting to Excel" & vbCrLf & _
                                        "  row: " & row_in_excel_where_to_put & ", column: " & column_in_excel_where_to_put
                                    WriteToExcel row_in_excel_where_to_put, column_in_excel_where_to_put, value, False
                                End If
                            Next
                        Else
                            row_factor_to_adust_for_skip = row_factor_to_adust_for_skip - 1
                            Log "Skipping SAP row " & r & " due to KTEXT containing 'Indutec'"
                        End If
                    Next
                    Log "Finished copying upcoming PMs to shift logbook."
                End If
                CleanUpAndTerminate "Finished script execution after detecting " & emptyRowsUntilDone & " empty rows."
            End If
        Else
            emptyRowCounter = 0
        End If


        If SkipReason = "" Then
            Log vbCrLf & "Checking row: " & i & " with the following content:" & vbCrLf & rowText
            If Check_if_WO_needs_confirmation(WO_Nr) Then
                Dim alleMitarbeiter, einzelnerMitarbeiter, confirmation, counter, total
                alleMitarbeiter = Split(Mitarbeiter, "/")
                total = UBound(alleMitarbeiter) + 1 'Ubound returns a 0-based index since arrays in VBScript are zero-based.
                counter = 1
                For Each einzelnerMitarbeiter In alleMitarbeiter
                    If counter = total Then
                        Log "confirmation counter: " & counter & " = total confirmations: " & total
                        confirmation = Confirm_WO(WO_Nr, einzelnerMitarbeiter, DauerInH, StartzeitConvertToUTC, EndzeitConvertToUTC, Massnahme, Status = done_text_from_excel)
                        ' Only write to excel if script is done with this row
                        WriteToExcel i, column_in_excel_where_to_put_message, confirmation, False
                    Else
                        Log "confirmation counter: " & counter & " // total confirmations: " & total
                        confirmation = Confirm_WO(WO_Nr, einzelnerMitarbeiter, DauerInH, StartzeitConvertToUTC, EndzeitConvertToUTC, Massnahme, False)
                    End If
                    counter = counter + 1
                Next
                'Else
                '    Log "WO " & WO_Nr & " does not need confirmation."
            End If
        Else
            Log "Skipping row " & i & " for the following reason(s): " & SkipReason
        End If
    End If
Next

' Performance improvement - problem: when script faults, not all is written! :(
' --- Flush buffered writes back to Excel in a single COM call when possible
' On Error Resume Next
' If lastRow >= 2 And IsArray(g_statusBuffer) Then
'     Dim outArr, r, rowsCount
'     rowsCount = lastRow - 1 ' rows from 2..lastRow
'     ReDim outArr(rowsCount - 1, 0)
'     For r = 2 To lastRow
'         outArr(r - 2, 0) = g_statusBuffer(r)
'     Next
'     ' Write the buffered column to Excel in one shot
'     sheet1.Range(sheet1.Cells(2, 16), sheet1.Cells(lastRow, 16)).Value = outArr
' End If
' On Error GoTo 0

' ' Restore Excel UI & calculation settings
' On Error Resume Next
' excelApp.ScreenUpdating = prevScreenUpdating
' If Not IsEmpty(prevCalculation) Then excelApp.Calculation = prevCalculation
' excelApp.EnableEvents = True
' On Error GoTo 0

CleanupAndTerminate "Finished."



Sub initialize()
    Dim logFolder, scriptName, userName, dateTimeStamp
    
    column_in_excel_where_to_put_message = 16
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get script name without extension
    scriptName = fso.GetBaseName(WScript.ScriptFullName)
    
    ' Get user name
    userName = CreateObject("WScript.Network").UserName
    
    ' Format date and time for filename
    dateTimeStamp = Year(Now) & "_" & Right("0" & Month(Now), 2) & "_" & Right("0" & Day(Now), 2) & "-" & _
        Right("0" & Hour(Now), 2) & "_" & Right("0" & Minute(Now), 2)
    
    ' Create the log folder path relative to the EXE location
    logFolder = fso.GetParentFolderName(WScript.ScriptFullName) & "\logs"
    
    If Not fso.FolderExists(logFolder) Then
        fso.CreateFolder(logFolder)
    End If
    
    ' Set the global log file path
    g_logFilePath = logFolder & "\" & scriptName & "_" & userName & "_" & dateTimeStamp & ".log"
    
    Set logFile = fso.OpenTextFile(g_logFilePath, 8, True) ' 8 = ForAppending
    Log "Initialized."
    Log "WScript.ScriptFullName: " & WScript.ScriptFullName
    Log "LogFilePath: " & g_logFilePath
    
    
    argFilePath = GetArgValue("filePath")
    argAutoConfirm = GetArgValue("autoConfirm")
    argUseCurrentExcel = GetArgValue("useCurrentExcel")
    
    ' Cache timezone bias once to avoid repeated registry reads during processing
    On Error Resume Next
    Dim sh
    Set sh = CreateObject("WScript.Shell")
    g_timezoneBias = ""
    g_timezoneBias = sh.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
    If Err.Number <> 0 Then
        Err.Clear
        g_timezoneBias = ""
        CleanupAndTerminate "ERROR caching active timezone bias (minutes to get to UTC from local time) from registry (HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias): " & Err.Number
    Else
        Log "Caching active timezone bias (minutes to get to UTC from local time) from registry (HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias): " & g_timezoneBias
    End If
    On Error GoTo 0
    
    ' === OPEN EXCEL ===
    On Error Resume Next
    Log "Trying to get existing Excel instance"
    Set excelApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Log "Excel not already open. Starting new instance..."
        Err.Clear
        Set excelApp = CreateObject("Excel.Application")
    Else
        Log "Excel app is already open. Using existing Excel-instance, which has " & excelApp.Workbooks.Count & " workbooks open already..."
    End If
    excelApp.DisplayAlerts = True
    excelApp.ScreenUpdating = True
    excelApp.Visible = True
    On Error GoTo 0
    
    
    If argFilePath <> "" Then ' take excel filepath from argument passed to the script
        filePath = argFilePath
        Log "using filePath from argument. argFilePath=" & argFilePath
    ElseIf argUseCurrentExcel = "yes" Then
        Log "using current excel because argument useCurrentExcel=" & argUseCurrentExcel
        Dim monthNamesGerman, currentMonthGerman
        monthNamesGerman = Array("Januar", "Februar", "März", "April", "Mai", "Juni", _
            "Juli", "August", "September", "Oktober", "November", "Dezember")
        currentMonthGerman = monthNamesGerman(Month(Date) - 1)
        filePath = "https://cargillonline.sharepoint.com/sites/mrdokumentenmanagement-2do/shared documents/2do/5 schichtbuch/" & currentMonthGerman & "_" & Year(Date) & "_Schichtbuch.xlsx"
        Log "using current excel path=" & filePath
    Else ' 
        Log "Showing file dialog so user can select Excel file..."
        filePath = excelApp.GetOpenFilename( _
            "Excel files (*.xlsx;*.xlsm;*.xls;*.xlsb;*.csv),*.xlsx;*.xlsm;*.xls;*.xlsb;*.csv", _
            , "Select your shift logbook Excel file that the script should read and sync to SAP", , False)
        If filePath = False Then
            CleanupAndTerminate "User cancelled the open file dialog in excel where he had to select the shift logbook Excel file."
        End If
        Log "using user-selected excel path=" & filePath
    End If
    
    If excelApp.Workbooks.Count <> 0 Then
        Log "Checking if shift logbook excel file is already open in one of the existing " & excelApp.Workbooks.Count & " excel-instances..."
        Dim wb, workbookAlreadyOpen : workbookAlreadyOpen = False
        For Each wb In excelApp.Workbooks
            If LCase(wb.FullName) = LCase(filePath) Then
                Log "Shift logbook excel file is already open. Using the already opened shift logbook excel to avoid reopening it..."
                Set workbook = wb
                workbookAlreadyOpen = True
                Exit For
            End If
        Next
    End If
    
    If Not workbookAlreadyOpen Then
        On Error Resume Next
        Set workbook = excelApp.Workbooks.Open(filePath)
        If Err.Number <> 0 Or workbook Is Nothing Then
            Err.Clear
            CleanupAndTerminate "Fehler beim Öffnen der Datei: " & filePath & vbCrLf & _
                "Details: " & Err.Description
        End If
        Log "Opened: " & filePath
        On Error GoTo 0
    End If
    Set sheet1 = workbook.sheets(1)
    Set sheet2 = workbook.sheets(2)
    Log "Shift logbook Excel file has " & workbook.sheets.Count & " sheets. sheet1 name: " & sheet1.Name & " // sheet2 name: " & sheet2.Name
    
    
    ' === FIND LAST ROW AND COLUMN ===
    lastRow = sheet1.Cells(sheet1.Rows.Count, 1).End( - 4162).Row ' -4162 = xlUp
    lastCol = sheet1.Cells(1, sheet1.Columns.Count).End( - 4159).Column ' -4159 = xlToLeft
    'Log "lastRow in Excel workbook:" & lastRow
    
    ' --- disable screen updating and calculation while processing to speed up Excel operations
    'On Error Resume Next
    'prevScreenUpdating = excelApp.ScreenUpdating
    'prevCalculation = excelApp.Calculation
    'excelApp.ScreenUpdating = False
    'On Error GoTo 0
    
    ' --- Performance improvement: cache Excel data once to variant 2D array
    ' Read the used range into memory (single COM call)
    sheet1_cached = Empty
    On Error Resume Next
    sheet1_cached = sheet1.Range(sheet1.Cells(1, 1), sheet1.Cells(lastRow, lastCol)).Value
    If Err.Number <> 0 Then CleanupAndTerminate "=== ERROR === Bulk read of sheet1 range failed."
    sheet2_cached = sheet2.Range(sheet2.Cells(1, 1), sheet2.Cells(lastRow, lastCol)).Value
    If Err.Number <> 0 Then CleanupAndTerminate "=== ERROR === Bulk read of sheet2 range failed."
    On Error GoTo 0
    
    done_text_from_excel = sheet2_cached(4, 5) ' cell E4 in sheet2
    cancelled_text_from_excel = sheet2_cached(5, 5) ' cell E5 in sheet2
    Log "done_text_from_excel sheet2: '" & done_text_from_excel & "' // cancelled_text_from_excel sheet2: '" & cancelled_text_from_excel & "'"
    
    Log vbCrLf & "Logging in to SAP..."
    loadedFromMainScript = True ' flag to indicate this script is calling the login script
    Set file = fso.OpenTextFile(fso.GetParentFolderName(WScript.ScriptFullName) & "\SAP Login.vbs", 1)
    code = file.ReadAll
    file.Close
    ExecuteGlobal code
    Set session = SAPLogin()
    
    If argAutoConfirm = "yes" Then
        autoConfirmResponse = vbYes
    ElseIf argAutoConfirm = "no" Then
        autoConfirmResponse = vbNo
    Else
        autoConfirmResponse = MsgBox(vbCrLf & _
            "Yes:" & vbCrLf & _
            "will confirm & close work orders automatically." & vbCrLf & vbCrLf & _
            "No:" & vbCrLf & _
            "will ask you before confirming & closing every single time.", vbYesNo + vbQuestion, _
            "Do you want to confirm and close WOs automatically in SAP?")
    End If
End Sub

Function Check_if_WO_needs_confirmation(wo_Nr)
    SafeStartTransaction "IW32"
    SafeSetText "wnd[0]/usr/ctxtCAUFVD-AUFNR", wo_Nr
    SafeSendVKey "wnd[0]", 0
    Dim returnValueFromSAP : returnValueFromSAP = SafeGetText("wnd[0]/sbar") ' attach return value from SAP confirmation number
    If returnValueFromSAP <> "" Then
        Log "SAP returned: " & returnValueFromSAP
        WriteToExcel i, column_in_excel_where_to_put_message, returnValueFromSAP, False
        Exit Function
    End If
    

    Dim sysStatus : sysStatus = GetSysStatus()
    If InStr(sysStatus, "TECO") > 0 Then
        Log "WO " & wo_Nr & " already completed."
        WriteToExcel i, column_in_excel_where_to_put_message, "WO " & wo_Nr & " already completed.", False
        Exit Function
    ElseIf InStr(sysStatus, "CLSD") > 0 Then
        Log "WO " & wo_Nr & " already closed."
        WriteToExcel i, column_in_excel_where_to_put_message, "WO " & wo_Nr & " already closed.", False
        Exit Function
    ElseIf InStr(sysStatus, "CNF") > 0 And InStr(sysStatus, "PCNF") = 0 Then ' not completed, not closed, confirmed (CNF), not partially confirmed (PCNF)
        Log "WO " & wo_Nr & " not completed/closed. Confirmed (CNF), but not partially confirmed (PCNF). Checking whether to complete WO..."
        Dim needs_TECO : needs_TECO = Check_if_WO_needs_TECO(wo_Nr)
        If needs_TECO <> "" Then WriteToExcel i, column_in_excel_where_to_put_message, needs_TECO, False
        Exit Function
    End If
    If Check_if_WO_contains_skip_condition() Then Exit Function
    If InStr(sysStatus, "REL ") > 0 And InStr(sysStatus, "RELR") = 0 Then ' released (REL) & NOT release rejected (RELR) & NOT technically completed (not TECO) & NOT Confirmed (not CNF)
        Log "WO " & wo_Nr & " released (REL) & NOT technically completed (not TECO) & NOT Confirmed (not CNF). WO needs confirmation..."
    Else
        Log "WO " & wo_Nr & " NOT yet released (but also neither confirmed nor complete). Releasing WO now... WO needs confirmation after that..."
        SafeSendVKey "wnd[0]", 25 ' CTRL+F1 releases the order
        SafeSendVKey "wnd[0]", 11 ' CTRL+S saves the order
    End If
    Check_if_WO_needs_confirmation = True
End Function

Function Confirm_WO(wo, mitarbeiter, dauerInStunden, startzeit, endzeit, massnahme, finalConfirmation)
    Log "Confirming WO      : " & wo & vbCrLf & _
        " Employee          : " & mitarbeiter & vbCrLf & _
        " Duration (hours)  : " & dauerInStunden & vbCrLf & _
        " Start (UTC)       : " & startzeit & vbCrLf & _
        " End (UTC)         : " & endzeit & vbCrLf & _
        " What was done     : " & massnahme & vbCrLf & _
        " finalConfirmation?  " & finalConfirmation
    
    SafeStartTransaction "IW41"
    ' enable Parameters > Goods movements > all components - see https://cargillonline.sharepoint.com/:i:/r/sites/SAPSUBERLIN/Shared%20Documents/Allgemeines/Anleitungen/Anwendung%20-%20Maintenance%20%26%20Reliability/Korrektive%20Instandhaltung%20(Currative%20Maintenance)/6_Arbeitszeitbest%C3%A4tigung%20(Time%20conformation)/Materialaustrag_Parameter-Einstellung.png
    SafeSendVKey "wnd[0]", 18 ' opens Parameters
    SafeSetSelected "wnd[1]/usr/chkTCORU-ACOMP", True ' ticks Goods movements > all components
    SafeSendVKey "wnd[1]", 0 ' enter
    
    SafeSetText "wnd[0]/usr/ctxtCORUF-AUFNR", wo
    ' Set operation number from excel column Q
    Dim planned_operation_from_excel : planned_operation_from_excel = GetCellFromCache(sheet1_cached, i, 17) ' column Q = 17
    If planned_operation_from_excel = "" Then
        Log "No planned operation provided in excel for logging work on. Assuming work relates to the first operation (0010)."
        SafeFindById("wnd[0]/usr/txtCORUF-VORNR").text = "0010"
    Else
        Log "Setting operation number from excel column Q: " & planned_operation_from_excel
        SafeFindById("wnd[0]/usr/txtCORUF-VORNR").text = planned_operation_from_excel ' column Q = 17
    End If

    SafeSendVKey "wnd[0]", 0 'enter
    
    
    Dim personnelNo, duration
    On Error Resume Next
    personnelNo = GetpersonnelNumber(mitarbeiter)
    If Err.Number <> 0 Then
        Confirm_WO = "ERROR: " & Err.Description & vbCrLf & "Personell no. field not found by ID. WO: '" & wo_Nr & "' might have >1 operations planned with different work centers - example WO 416061308. Suggesting you manually do the confirmation."
        Log "Confirm_WO(" & wo_Nr & "): " & vbCrLf & Confirm_WO
        'Dim errorResponse
        'errorResponse = MsgBox(Confirm_WO, vbOKCancel + vbQuestion)
        'If errorResponse = vbCancel Then
        '    CleanupAndTerminate "user clicked cancel. Error was: " & Confirm_WO
        'End If
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0 ' Turn back on default error handling
    If Len(personnelNo) <> 8 Then
        Confirm_WO = "ERROR: personellNo: '" & personnelNo & "' is not 8 chars long."
    End If
    
    SafeSetText "wnd[0]/usr/ctxtAFRUD-PERNR", personnelNo
    duration = ConvertTimeToDecimalHour(dauerInStunden)
    SafeSetText "wnd[0]/usr/txtAFRUD-ISMNW_2", duration ' Actual work
    SafeSetText "wnd[0]/usr/txtAFRUD-IDAUR", TrimAfterComma(duration) ' Actual duration - Actual work can be X,YZ while Actual duration can only be one char after comma e.g. X,Y
    SafeSetText "wnd[0]/usr/ctxtAFRUD-ISDD", DateValue(startzeit) ' work start day
    SafeSetText "wnd[0]/usr/ctxtAFRUD-ISDZ", TimeValue(startzeit) ' work start time
    SafeSetText "wnd[0]/usr/ctxtAFRUD-IEDD", DateValue(endzeit) ' work end day
    SafeSetText "wnd[0]/usr/ctxtAFRUD-IEDZ", TimeValue(endzeit) ' work end time
    If Len(massnahme) > 40 Then
        Log "'Massnahme' has more than 40 chars, but SAP confirmation text cuts off after 40 chars. Long text input not yet programmed so shortening massnahme to first 40 chars"
        massnahme = Left(massnahme, 40)
    End If
    SafeSetText "wnd[0]/usr/txtAFRUD-LTXA1", massnahme ' confirm text short - 40 char max
    SafeSendVKey "wnd[0]", 0 ' enter so user can see name next to personnel number
    If SafeGetText("wnd[0]/sbar") <> "" Then
        CleanupAndTerminate "Unable to confirm WO - " & SafeGetText("wnd[0]/sbar")
    End If
    
    If finalConfirmation Then
        SafeSetText "wnd[0]/usr/txtAFRUD-OFMNW_2", "0" ' Remaining work
        SafeSetSelected "wnd[0]/usr/chkAFRUD-AUERU", True ' Final confirmation
        SafeSetSelected "wnd[0]/usr/chkAFRUD-LEKNW", True ' No remaining work
    End If
    
    Dim ConfirmationResponse
    If autoConfirmResponse = vbYes Then
        ConfirmationResponse = vbYes
    Else
        ConfirmationResponse = MsgBox("Do you want to confirm WO " & wo & " in SAP?" & vbCrLf & vbCrLf & _
            "Yes:" & vbCrLf & _
            "will send the confirmation as you see it on your screen and proceed with the script." & vbCrLf & vbCrLf & _
            "No:" & vbCrLf & _
            "will NOT send the confirmation but proceed with the script" & vbCrLf & vbCrLf & _
            "Cancel:" & vbCrLf & _
            "will NOT send the confirmation and TERMINATE the script.", vbYesNoCancel + vbQuestion, "Entry for WO " & wo & " found in Excel")
    End If
    Select Case ConfirmationResponse
        Case vbYes
            'Confirm_WO = SafeFindById("wnd[0]/usr/txtAFVGD-RUECK").text ' save the confirmation number - must do so before enter or save, otherwise this number might not be accesible bc we're in another menu
            SafeSendVKey "wnd[0]", 8 'F8 to open Goods movements overview
            SafeSendVKey "wnd[0]", 11 ' Saves the confirmation OR confirmation with goods movement
            'Confirm_WO = Confirm_WO & " // " & 'SafeFindById("wnd[0]/sbar").text ' attach return value from SAP confirmation number e.g. "Number of confirmations saved for order 417052748: 1"
            Confirm_WO = "WO confirmed"

            'TODO: might get a warning from SAP here that ticking 'no remaining work' will set the remaining work field to zero. 
            'This happens e.g. when a WO is planned with 1hr, but only 0,5h was worked for it (e.g. 416100298). TODO guard against this
            If finalConfirmation Then
                WScript.Sleep 500 ' hate to be doing this, but SAP can still be working on this work order after the script confirmed it (i.e. when there are goods movements), so waiting here to not error out bc work order is still being processed 
                Confirm_WO = Confirm_WO & " // " & Check_if_WO_needs_TECO(wo)
            End If
        Case vbNo
            Confirm_WO = "user clicked no when asked whether to confirm WO " & wo & ". Not confirming, but proceeding with the script..."
        Case vbCancel
            CleanupAndTerminate "user clicked cancel when asked whether to confirm WO " & wo & ". Not confirming and terminating the script."
    End Select
    Log "Confirm_WO " & wo_Nr & " returned: " & Confirm_WO
End Function

Function Check_if_WO_needs_TECO(wo_Nr)
    SafeStartTransaction "IW32"
    SafeSetText "wnd[0]/usr/ctxtCAUFVD-AUFNR", wo_Nr
    SafeSendVKey "wnd[0]", 0

    Dim sysStatus : sysStatus = GetSysStatus()
    If InStr(sysStatus, "TECO") > 0 Then
        Log "WO " & wo_Nr & " already completed."
        WriteToExcel i, column_in_excel_where_to_put_message, "WO " & wo_Nr & " already completed.", False
        Exit Function
    ElseIf InStr(sysStatus, "CLSD") > 0 Then
        Log "WO " & wo_Nr & " already closed."
        WriteToExcel i, column_in_excel_where_to_put_message, "WO " & wo_Nr & " already closed.", False
        Exit Function
    End If
    
    Dim CNF_Not_CAPR_response, orderText, objNetwork
    orderText = SafeGetText("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell")
    SafeSetText "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell", _
        orderText & vbCr & vbCr & _
        "Completed by script executed by user " & username & " on " & Now & " using data from excel" & vbCr & filePath
    
    If autoConfirmResponse = vbYes Then
        CNF_Not_CAPR_response = vbYes
    Else
        CNF_Not_CAPR_response = MsgBox("Do you want to complete WO " & wo_Nr & " in SAP?" & vbCrLf & vbCrLf & _
            "Yes:" & vbCrLf & _
            "will complete the WO as you see it on your screen and proceed with the script." & vbCrLf & vbCrLf & _
            "No:" & vbCrLf & _
            "will NOT complete the WO but proceed with the script" & vbCrLf & vbCrLf & _
            "Cancel:" & vbCrLf & _
            "will NOT complete the WO and TERMINATE the script.", vbYesNoCancel + vbQuestion, "Complete WO " & wo_Nr & " in SAP?")
    End If
    Select Case CNF_Not_CAPR_response
        Case vbYes
            SafeSendVKey "wnd[0]", 36 ' CTRL+F12
            SafeSendVKey "wnd[0]", 0 ' Enter

            Dim msgAfterTECO
            If autoConfirmResponse = vbYes Then
                msgAfterTECO = "WO TECO'd."
            Else
                msgAfterTECO = "WO " & wo_Nr & " TECO'd because user requested to."
            End If
            Log msgAfterTECO
            Check_if_WO_needs_TECO = msgAfterTECO
        Case vbNo
            Log "WO " & wo_Nr & " not completed. User clicked no when asked to complete, so I did not complete. Proceeding with the script..."
        Case vbCancel
            CleanupAndTerminate "WO " & wo_Nr & " not completed. User clicked cancel, so terminating script."
    End Select
End Function

' Gets the System-Status of a WO. 
' Prerequisite: Need to have work order open in SAP (Tcode IW32 for example)
Function GetSysStatus()
    Dim sysStatusObject
    On Error Resume Next
    Set sysStatusObject = SafeFindById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-STTXT")
    If Err.Number <> 0 Then
        Log "sysStatusObject not found by ID for WO " & wo_Nr & ". Error: " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0 ' Turn back on default error handling

    GetSysStatus = SafeGetText("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-STTXT")
    Log "sysStatus of WO : " & GetSysStatus
End Function

Function GetpersonnelNumber(employeeName)
    Dim employeeName_Normalized
    employeeName_Normalized = NormalizeString(employeeName)
    
    ' Search in sheet2_cached for matching employee name and return SAP personnel number
    Dim r, employeeName_in_sheet2_Normalized, employeeName_in_sheet2, personnelNo

    If Not IsEmpty(sheet2_cached) And Not IsNull(sheet2_cached) Then
        ' Iterate through sheet2_cached rows (starting from row 2, assuming row 1 is header)
        For r = 2 To UBound(sheet2_cached, 1)
            employeeName_in_sheet2 = GetCellFromCache(sheet2_cached, r, 1) ' Get name from column 1

            ' Normalize the name from sheet2 the same way as the input
            employeeName_in_sheet2_Normalized = NormalizeString(employeeName_in_sheet2)
            If employeeName_in_sheet2_Normalized = "" Then Exit For

            ' Compare normalized names - allow lenient matching
            If CompareStringsLeniently(employeeName_Normalized, employeeName_in_sheet2_Normalized) Then
                personnelNo = GetCellFromCache(sheet2_cached, r, 2) ' Get SAP personnel number from column 2 (=column B)
                If personnelNo <> "" Then
                    Log "Employee '" & employeeName & "' found in sheet2, row " & r & " as '" & employeeName_in_sheet2 & "', mapped to personnel number from sheet2: " & personnelNo
                    GetpersonnelNumber = personnelNo
                    Exit Function
                Else
                    CleanupAndTerminate "Employee '" & employeeName & "' found in sheet2, row " & r & " as '" & employeeName_in_sheet2 & "', but personnel number field is empty."
                End If
            End If
        Next
    End If

    ' If no match found in sheet2, log and return empty (don't confirm with personnel number)
    Log vbCrLf & "=== WARNING === Unknown employee: " & employeeName & " // normalized to: " & employeeName_Normalized & " // Not found in sheet2. Confirming without entering employee to SAP..."
End Function

' Normalizes a string for lenient comparison: lowercases, removes accents, removes spaces, removes doubled letters
Function NormalizeString(inputStr)
    NormalizeString = DeleteNewLineAndTrim(inputStr)
    NormalizeString = LCase(NormalizeString) ' ignore capitalization
    NormalizeString = Replace(NormalizeString, " ", "") ' remove spaces
    NormalizeString = Replace(NormalizeString, "-", "") ' remove hyphens
    NormalizeString = Replace(NormalizeString, "ä", "ae")
    NormalizeString = Replace(NormalizeString, "ö", "oe")
    NormalizeString = Replace(NormalizeString, "ü", "ue")
    NormalizeString = Replace(NormalizeString, "ß", "s")
    NormalizeString = Replace(NormalizeString, "tzt", "tz")
    NormalizeString = Replace(NormalizeString, "zt", "tz")
    NormalizeString = Replace(NormalizeString, "schp", "sp")
    NormalizeString = RemoveDoubleLetters(NormalizeString)
    NormalizeString = RemoveDoubleLetters(NormalizeString)
End Function

' Compares two strings leniently, returns True not only for exact matches but also with minimal character differences
Function CompareStringsLeniently(string1, string2)
    CompareStringsLeniently = False
    
    'Log "Comparing strings leniently: '" & string1 & "' vs. '" & string2 & "'"

    ' Exact match
    If string1 = string2 Then
        CompareStringsLeniently = True
        Exit Function
    End If

    ' Only compare if lengths are reasonably similar (within 2 characters)
    Dim len1, len2, lengthDiff, maxLen, charDiff
    len1 = Len(string1)
    len2 = Len(string2)

    If len1 > len2 Then
        lengthDiff = len1 - len2
        maxLen = len1
    Else
        lengthDiff = len2 - len1
        maxLen = len2
    End If

    ' Reject if length difference is too large (e.g., "nicostuijt" is 10 chars, "michelgroen" is 11 chars - too different)
    If lengthDiff > 2 Then
        Exit Function
    End If

    ' Only attempt lenient matching for reasonably long names
    If maxLen < 5 Then
        Exit Function
    End If

    ' Count character differences by comparing position by position
    charDiff = 0
    Dim i, minLen
    minLen = Len(string1)
    If minLen > Len(string2) Then
        minLen = Len(string2)
    End If

    For i = 1 To minLen
        If Mid(string1, i, 1) <> Mid(string2, i, 1) Then
            charDiff = charDiff + 1
        End If
    Next

    ' Add the length difference to character difference
    charDiff = charDiff + lengthDiff

    ' Allow match only if character differences are minimal (1-2 for short names, up to 3 for longer names)
    Dim charDiffThreshold
    If maxLen < 7 Then
        charDiffThreshold = 1
    Else
        charDiffThreshold = 2
    End If

    If charDiff <= charDiffThreshold Then
        CompareStringsLeniently = True
    End If
End Function

Function RemoveDoubleLetters(inputStr)
    Dim i, letter
    For i = 65 To 90 ' ASCII A-Z
        letter = Chr(i)
        inputStr = Replace(inputStr, letter & letter, letter)
        inputStr = Replace(inputStr, LCase(letter) & LCase(letter), LCase(letter))
    Next
    RemoveDoubleLetters = inputStr
End Function

Function DeleteNewLineAndTrim(inputString)
    inputString = Replace(inputString, vbCrLf, "")
    inputString = Replace(inputString, vbCr, "")
    inputString = Replace(inputString, vbLf, "")
    inputString = Trim(inputString) ' ignore leading & lagging spaces
    DeleteNewLineAndTrim = inputString
End Function

' Converts Time like 01:30:00 (1 hour 30 minutes) to decimal hour like 1,5 hrs for SAP input
Function ConvertTimeToDecimalHour(timeInput)
    Dim parts, hours, minutes, seconds
    
    parts = Split(timeInput, ":")
    
    hours = CInt(parts(0))
    minutes = CInt(parts(1))
    seconds = CInt(parts(2))
    
    ConvertTimeToDecimalHour = hours + (minutes / 60) + (seconds / 3600)
    
    ' Replace dot with comma for format settings
    ConvertTimeToDecimalHour = Replace(FormatNumber(ConvertTimeToDecimalHour, 2), ".", ",")
End Function

'In Excel, time values are stored as fractions of a day. So when you extract a time like 0.416666666666667 from Excel, it actually represents: 0.416666666666667 × 24 = 10 hours - This function then returns 10:00:00.
Function ConvertExcelFractionToTime(excelTime)
    ' Handle null or empty input
    If IsNull(excelTime) Or Trim(CStr(excelTime)) = "" Then
        Log "ExcelFractionToTime(" & excelTime & ") - parameter is null or empty"
        ConvertExcelFractionToTime = Null
        Exit Function
    End If
    
    ' Check if it's a valid numeric value
    If Not IsNumeric(excelTime) Then
        Log "ExcelFractionToTime(" & excelTime & ") - parameter is not numeric"
        ConvertExcelFractionToTime = Null
        Exit Function
    End If
    
    ' Convert to time if it's a fraction between 0 and 1
    If CDbl(excelTime) >= 0 And CDbl(excelTime) < 1 Then


        Dim totalSeconds, hours, minutes, seconds

        ' Round total seconds to avoid truncation errors
        totalSeconds = Round(excelTime * 86400)

        ' Calculate hours, minutes, seconds
        hours = totalSeconds \ 3600
        minutes = (totalSeconds Mod 3600) \ 60
        seconds = totalSeconds Mod 60

        ConvertExcelFractionToTime = TimeSerial(hours, minutes, seconds)

    Else
        Log "ExcelFractionToTime(" & excelTime & ") - parameter is not between 0 and 1"
        ConvertExcelFractionToTime = Null
    End If
End Function

' Converts DateTime like "21.08.2025 01:30:00" safely to Date format. Returns Null if conversion fails
Function SafeCDate(value)
    Dim result
    
    ' Check for empty or null input
    If IsNull(value) Or Trim(CStr(value)) = "" Then
        Log "SafeCDate(" & value & ") - parameter is null or empty"
        SafeCDate = Null
        Exit Function
    End If
    
    ' Check if it's numeric or date
    If Not IsNumeric(value) And Not IsDate(value) Then
        Log "SafeCDate(" & value & ") - parameter is not numeric and not a date"
        SafeCDate = Null
        Exit Function
    End If
    
    ' Try converting full date-time string
    On Error Resume Next
    SafeCDate = CDate(value)
    If Err.Number <> 0 Then
        Log "Error when attempting to CDate(" & value & ") : " & Err.Description
        SafeCDate = Null
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Safely access a cell. Avoids faulting the script when a cell contains an image for example. If access error: Echoes the error and returns an empty string
Function SafeCellAccess(sheet, row, column)
    If row < 1 Or column < 1 Then
        CleanupAndTerminate "ERRROR: Function SafeCellAccess(sheetName=" & sheet.Name & " & row=" & row & ", column=" & column & ") failed because row or column or both are < 1."
    End If
    
    On Error Resume Next
    Dim val
    val = sheet.Cells(row, column).Value
    If Err.Number <> 0 Then
        Log "SafeCellAccess(sheetName=" & sheet.Name & " & row=" & row & ", column=" & column & ") failed with error: " & Err.Description & ". Returning an empty string."
        SafeCellAccess = ""
        Err.Clear
    Else
        If IsEmpty(val) Or IsNull(val) Then
            SafeCellAccess = ""
        Else
            SafeCellAccess = CStr(val)
            SafeCellAccess = DeleteNewLineAndTrim(SafeCellAccess)
        End If
    End If
    On Error GoTo 0
End Function

' In SAP, Actual work can be X,YZ while Actual duration can only be one char after separator e.g. X,Y
Function TrimAfterComma(inputStr)
    ' Example usage:
    'Dim original, trimmed
    'original = "123,456"
    'trimmed = TrimAfterComma(original)
    'MsgBox trimmed  ' Output: 123,4
    Dim parts, result
    parts = Split(inputStr, ",")
    
    If UBound(parts) = 0 Then
        result = inputStr ' No comma found
    Else
        result = parts(0) & "," & Left(parts(1), 1)
    End If
    
    TrimAfterComma = result
End Function

Function ConvertToUTC(inputDate)
    On Error Resume Next
    
    Dim bias, utcDate
    ' Use cached timezone bias if available (set in initialize())
    If Not IsEmpty(g_timezoneBias) And g_timezoneBias <> "" Then
        bias = g_timezoneBias
    Else
        Dim objShell
        Set objShell = CreateObject("WScript.Shell")
        bias = objShell.RegRead("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\TimeZoneInformation\\ActiveTimeBias")
        If Err.Number <> 0 Then
            Log "ConvertToUTC(" & inputDate & ") - Error reading time zone bias: " & Err.Description
            ConvertToUTC = inputDate ' Return input as fallback
            Err.Clear
            Exit Function
        End If
    End If
    
    ' Validate inputTime
    If Not IsDate(inputDate) Then
        Log "ConvertToUTC(" & inputDate & ") - Invalid input: not a valid date/time."
        ConvertToUTC = inputDate
        Exit Function
    End If
    
    ' Convert to UTC using minutes bias
    utcDate = DateAdd("n", bias, inputDate)
    If Err.Number <> 0 Then
        Log "ConvertToUTC(" & inputDate & ") - Error during DateAdd: " & Err.Description
        ConvertToUTC = inputDate
        Err.Clear
        Exit Function
    End If
    
    ConvertToUTC = utcDate
    On Error GoTo 0
End Function

Function IsInArray(valToCheck, arr)
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = valToCheck Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

' Writes a message to the specified cell in Excel sheet1.
' Validates parameters and checks if the sheet is protected before writing.
Sub WriteToExcel(row, column, message, redBackgroundAndBoldText)
    Log "Attempting to write to Excel:" & vbCrLf & _
        "  Row: " & row & vbCrLf & _
        "  Column: " & column & vbCrLf & _
        "  Message: " & message & vbCrLf & _
        "  RedBackgroundAndBoldText: " & redBackgroundAndBoldText
    If row < 1 Then
        CleanupAndTerminate "ERROR: Sub WriteToExcel(row=" & row & ", column=" & column & ", message=" & message & ", redBackgroundAndBoldText=" & redBackgroundAndBoldText & ") failed because row is < 1."
    End If
    If column < 1 Then
        CleanupAndTerminate "ERROR: Sub WriteToExcel(row=" & row & ", column=" & column & ", message=" & message & ", redBackgroundAndBoldText=" & redBackgroundAndBoldText & ") failed because column is < 1."
    End If
    If IsNull(message) Then
        CleanupAndTerminate "ERROR: Sub WriteToExcel(row=" & row & ", column=" & column & ", message=" & message & ", redBackgroundAndBoldText=" & redBackgroundAndBoldText & ") failed because message is null."
    End If
    If IsEmpty(sheet1) Or IsNull(sheet1) Then
        CleanupAndTerminate "ERROR: Sub WriteToExcel(row=" & row & ", column=" & column & ", message=" & message & ", redBackgroundAndBoldText=" & redBackgroundAndBoldText & ") failed because sheet1 is not initialized."
    End If
    
    On Error Resume Next
    ' Check if sheet1 is protected
    If sheet1.ProtectContents Or sheet1.ProtectDrawingObjects Or sheet1.ProtectScenarios Then
        CleanupAndTerminate "ERROR: Sub WriteToExcel(row=" & row & ", message=" & message & ", redBackgroundAndBoldText=" & redBackgroundAndBoldText & ") failed: sheet1 is protected."
    End If

    If Not IsNull(redBackgroundAndBoldText) And redBackgroundAndBoldText = True Then
        sheet1.Cells(row, column).Interior.Color = RGB(255, 0, 0) ' Red background for error messages
        sheet1.Cells(row, column).Font.Bold = True ' Bold font for error messages
    End If
    
    Dim existingValue : existingValue = SafeCellAccess(sheet1, row, column)
    If existingValue = "" Then
        sheet1.Cells(row, column).Value = message
    Else
        Log "WARNING! Not overwriting existing value in Excel at row " & row & ", column " & column & ". Existing value: '" & existingValue & "'"
    End If

    ' Check for error
    If Err.Number <> 0 Then
        Log "Write failed: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Return a cell value from the cached worksheet1 (g_data) when available, otherwise fall back to SafeCellAccess
Function GetCellFromCache(sheet_cached, r, c)
    On Error Resume Next
    Dim rawVal
    If IsEmpty(sheet_cached) Or IsNull(sheet_cached) Then
        rawVal = SafeCellAccess(sheet_cached, r, c)
    Else
        rawVal = sheet_cached(r, c)
        If Err.Number <> 0 Then
            Err.Clear
            rawVal = SafeCellAccess(sheet_cached, r, c)
        End If
    End If
    On Error GoTo 0
    
    ' Normalize output similar to previous SafeValue behaviour
    If IsNull(rawVal) Or IsEmpty(rawVal) Then
        GetCellFromCache = ""
        Exit Function
    End If
    If IsObject(rawVal) Then
        GetCellFromCache = ""
        Exit Function
    End If
    On Error Resume Next
    GetCellFromCache = CStr(rawVal)
    If Err.Number <> 0 Then
        Err.Clear
        GetCellFromCache = ""
    Else
        GetCellFromCache = DeleteNewLineAndTrim(GetCellFromCache)
    End If
    On Error GoTo 0
End Function

' Returns True if the entire cached row (1..lastCol) is empty or whitespace.
' Uses the in-memory `g_data` via `GetCellFromCache` and `SafeValue`.
' Falls back to cell-by-cell access when no cache is available.
Function IsCachedRowEmpty(rowIndex)
    Dim c, cellVal
    
    ' Validate row index
    If rowIndex < 1 Then
        IsCachedRowEmpty = True
        Exit Function
    End If
    
    ' If lastCol is not set, try to determine it conservatively
    If IsEmpty(lastCol) Or lastCol = 0 Then
        IsCachedRowEmpty = True
        Exit Function
    End If
    
    For c = 1 To lastCol
        cellVal = GetCellFromCache(sheet1_cached, rowIndex, c)
        If cellVal <> "" Then
            IsCachedRowEmpty = False
            Exit Function
        End If
    Next
    
    IsCachedRowEmpty = True
End Function

' Writes the message to both the console (WScript.Echo) and the log file.
Sub Log(message)
    Dim time_and_message : time_and_message = Now & " - " & message
    WScript.Echo time_and_message
    ' Append the message to the log file
    logFile.WriteLine time_and_message
End Sub

' Gets the value of an argument "argName" passed to the script at runtime
Function GetArgValue(argName)
    Dim i, arg, parts
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If InStr(arg, "=") > 0 Then
            parts = Split(arg, "=")
            If LCase(parts(0)) = LCase(argName) Then
                GetArgValue = parts(1)
                Log "Argument '" & argName & "' with value: '" & GetArgValue & "'"
                Exit Function
            End If
        End If
    Next
    GetArgValue = ""
End Function

' Returns a timestamp in format HH:MM:SS.mmm
Function Timestamp()
    Dim t
    t = Timer ' e.g., 45296.234
    
    Dim hours, minutes, seconds, milliseconds
    hours = Int(t \ 3600)
    minutes = Int((t Mod 3600) \ 60)
    seconds = Int(t Mod 60)
    milliseconds = Int((t - Int(t)) * 1000)
    timestamp = Right("0" & hours, 2) & ":" & Right("0" & minutes, 2) & ":" & Right("0" & seconds, 2) & "." & Right("00" & milliseconds, 3)
End Function

' Checks if a WO contains a skip condition. Returns False if no skip condition in WO, true if at least one skip condition in WO. 
' Prerequisite: Need to have work order open in SAP (Tcode IW32 for example)
' skip conditions are:
' any purchase or (more than one planned operation in SAP - so Operation 0020 - BUT NO PLANNED OPERATION PROVIDED IN EXCEL for logging work on)
Function Check_if_WO_contains_skip_condition()
    Check_if_WO_contains_skip_condition = False
    If SAP_plantcode = "" Then SAP_plantcode = SafeGetText("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:0154/txtCAUFVD-IWERK")

    Log "Checking if WO contains skip condition for the script..."
    Log "1. skip condition: checking if WO has more than one operation in SAP but no planned operation in shift logbook excel."
    Log "   Checking operations tab in SAP for more than one operation..."
    ' Select operations tab
    SafeSelect "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpVGUE"
    ' Check if operation 0020 exists by trying to read its description field
    Dim operation_2_description
    operation_2_description = SafeGetText("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-LTXA1[7,1]") ' 0-based [column, row] - column 7 = description, row 1 = operation 0020
    If operation_2_description <> "" Then
        Log "   Found operation 0020 with description: '" & operation_2_description & "' in WO."
        Log "   Next checking whether planned operation provided in excel for logging work on..."
        Dim planned_operation_from_excel : planned_operation_from_excel = GetCellFromCache(sheet1_cached, i, 17) ' column Q = 17
        If planned_operation_from_excel = "" Then
            Log "   No planned operation provided in excel for logging work on. Assuming work relates to the first operation (0010). Continuing script..."
            'Dim ERROR_MESSAGE_no_planned_operation_provided_in_excel
            'ERROR_MESSAGE_no_planned_operation_provided_in_excel = "WO has multiple operations in SAP, but no planned operation provided in excel for logging work on. Script ignoring this work order. Please provide planned operation."
            'Log "   " & ERROR_MESSAGE_no_planned_operation_provided_in_excel
            'WriteToExcel i, column_in_excel_where_to_put_message, ERROR_MESSAGE_no_planned_operation_provided_in_excel, True
            'WriteToExcel i, column_in_excel_where_to_put_message + 1, "", True ' make cell where operation is supposed to be entered also red background to show user what is missing
            'Check_if_WO_contains_skip_condition = True
            'Exit Function
        Else
            'Log "   Planned operation " & planned_operation_from_excel & " provided in excel for logging work on."
            Dim planned_operation_key_for_SAP : planned_operation_key_for_SAP = planned_operation_from_excel
            ' Adjust planned_operation_key_for_SAP, e.g. from "10" to "1"
            ' Remove exactly one trailing zero if present and length > 1
            If Len(planned_operation_key_for_SAP) > 1 And Right(planned_operation_key_for_SAP, 1) = "0" Then
                planned_operation_key_for_SAP = Left(planned_operation_key_for_SAP, Len(planned_operation_key_for_SAP) - 1)
            Else
                CleanupAndTerminate "CRITICAL ERROR: Planned operation key for SAP from excel '" & planned_operation_key_for_SAP & "' is not in expected format. Expected format: at least two digits with trailing zero for operations, e.g. '10' for operation 1 (shown as 0010 in Excel), '20' for operation 2, etc."
            End If
            'Log "   Planned operation key for SAP after adjustment: " & planned_operation_key_for_SAP
            ' check if planned_operation_from_excel matches has a planned operation in SAP WO by checking if there is short text
            Dim operation_short_text_in_SAP_WO : operation_short_text_in_SAP_WO = SafeGetText("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpVGUE/ssubSUB_AUFTRAG:SAPLCOVG:3010/tblSAPLCOVGTCTRL_3010/txtAFVGD-LTXA1[7," & planned_operation_key_for_SAP - 1 & "]") ' 0-based [column, row]
            If operation_short_text_in_SAP_WO = "" Then
                Dim ERROR_MESSAGE_planned_operation_from_excel_has_no_short_text_in_SAP_WO
                ERROR_MESSAGE_planned_operation_from_excel_has_no_short_text_in_SAP_WO = "Planned operation from excel '" & planned_operation_from_excel & "' has no matching operation short text in SAP WO."
                Log "   " & ERROR_MESSAGE_planned_operation_from_excel_has_no_short_text_in_SAP_WO
                WriteToExcel i, column_in_excel_where_to_put_message, ERROR_MESSAGE_planned_operation_from_excel_has_no_short_text_in_SAP_WO, True
                WriteToExcel i, column_in_excel_where_to_put_message + 1, "", True ' make cell with wrong operation also red background to show user what is wrong
                Check_if_WO_contains_skip_condition = True
                Exit Function
            Else
                Log "   Planned operation '" & planned_operation_from_excel & "' provided in excel has operation short text '" & operation_short_text_in_SAP_WO & "' in SAP WO. Continuing script..."
            End If
        End If

    End If
    Log "1. skip condition looks good. Next checking..."
    Log "2. skip condition: at least one purchase. Browsing document flow..."
    SafeSendVKey "wnd[0]", 35 'CTRL+F11 (Document Flow)
    
    Dim tree, index, key, itemText
    Set tree = SafeFindById("wnd[0]/usr/shell/shellcont[1]/shell[1]")
    
    index = 1
    Do
        key = Right(Space(11) & index, 11) ' Pad index to 11 characters
        'Log "key padded to 10 chars:" & key
        'selectItemParameter = key & ", &Hierarchy"
        'Log "selectItemParameter:" & selectItemParameter
        On Error Resume Next
        itemText = tree.getItemText(key, "&Hierarchy")
        If Err.Number <> 0 Then
            Log "   " & index & " => No skip condition found."
            '"Error at key '" & key & "'" & vbCrLf & _
            '"Error Number: " & Err.Number & vbCrLf & _
            '"Description: " & Err.Description
            Exit Do
        End If
        On Error GoTo 0

        Log "   " & index & " => " & itemText

        If InStr(itemText, "Purchase") > 0 Then
            Dim NotificationMessage : NotificationMessage = "Found: " & itemText & vbCrLf & "in work order, therefore script ignoring this work order."
            WriteToExcel i, column_in_excel_where_to_put_message, NotificationMessage, True
            Log NotificationMessage
            Check_if_WO_contains_skip_condition = True
            Exit Do
        End If

        index = index + 1
    Loop
    SafeSendVKey "wnd[0]", 3 ' F3 to go back from document flow to WO in SAP GUI
    SafeSelect "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1101/tabsTS_1100/tabpIHKZ" ' Select HeaderData tab again to get back to initial state
End Function

' Safe wrapper to access SAP session objects. Returns the object if found, or terminates gracefully if not.
' Usage: Set myControl = SafeFindById("wnd[0]/usr/ctxtFieldName")
Function SafeFindById(objectPath)
    On Error Resume Next
    Dim obj
    Set obj = session.findById(objectPath)
    
    If Err.Number <> 0 Or obj Is Nothing Then
        Dim errorMsg
        errorMsg = "CRITICAL ERROR: SAP session object not found!" & vbCrLf & vbCrLf & _
            "Path: " & objectPath & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
            "This may indicate:" & vbCrLf & _
            "- SAP GUI is not responding or session ended" & vbCrLf & _
            "- The screen layout is different than expected" & vbCrLf & _
            "- A previous operation did not complete properly" & vbCrLf & vbCrLf & _
            "Terminating script..."
        CleanupAndTerminate errorMsg
    End If
    
    On Error GoTo 0
    Set SafeFindById = obj
End Function

' Safely start an SAP transaction. Terminates script on error with clear message.
' Usage: SafeStartTransaction "IW32"
Sub SafeStartTransaction(transactionCode)
    On Error Resume Next
    session.StartTransaction transactionCode
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = "CRITICAL ERROR: Unable to start SAP transaction '" & transactionCode & "'" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
            "This may indicate the SAP session ended or SAP GUI is not responding. Terminating script..."
        CleanupAndTerminate errMsg
    End If
    On Error GoTo 0
End Sub

' Safely send a virtual key to an SAP window. Terminates on error with a clear message.
' Usage: SafeSendVKey "wnd[0]", 0
Sub SafeSendVKey(windowPath, key)
    On Error Resume Next
    Dim wnd
    Set wnd = SafeFindById(windowPath)
    wnd.sendVKey key
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = "CRITICAL ERROR: sendVKey failed for '" & windowPath & "' with key: " & key & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
            "This may indicate the SAP GUI is not responding or the session ended. Terminating script..."
        CleanupAndTerminate errMsg
    End If
    On Error GoTo 0
End Sub

' Safely get the .text property of an SAP control. Terminates on error.
Function SafeGetText(objectPath)
    On Error Resume Next
    Dim obj, val
    Set obj = SafeFindById(objectPath)
    val = obj.text
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = "CRITICAL ERROR: Unable to read .text from '" & objectPath & "'" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
            "This may indicate the SAP GUI is not responding or the session ended. Terminating script..."
        CleanupAndTerminate errMsg
    End If
    On Error GoTo 0
    SafeGetText = val
End Function

' Safely set the .text property of an SAP control. Terminates on error.
Sub SafeSetText(objectPath, value)
    On Error Resume Next
    Dim obj
    Set obj = SafeFindById(objectPath)
    obj.text = value
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = "CRITICAL ERROR: Unable to set .text on '" & objectPath & "' to '" & value & "'" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
            "This may indicate the SAP GUI is not responding or the session ended. Terminating script..."
        CleanupAndTerminate errMsg
    End If
    On Error GoTo 0
End Sub

' Safely set the .selected property (e.g., checkboxes). Terminates on error.
Sub SafeSetSelected(objectPath, boolValue)
    On Error Resume Next
    Dim obj
    Set obj = SafeFindById(objectPath)
    obj.selected = boolValue
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = "CRITICAL ERROR: Unable to set .selected on '" & objectPath & "' to '" & boolValue & "'" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
            "This may indicate the SAP GUI is not responding or the session ended. Terminating script..."
        CleanupAndTerminate errMsg
    End If
    On Error GoTo 0
End Sub

' Safely call .press on a control. Terminates on error.
Sub SafePress(objectPath)
    On Error Resume Next
    Dim obj
    Set obj = SafeFindById(objectPath)
    obj.press
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = "CRITICAL ERROR: Unable to press '" & objectPath & "'" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
            "This may indicate the SAP GUI is not responding or the session ended. Terminating script..."
        CleanupAndTerminate errMsg
    End If
    On Error GoTo 0
End Sub

' Safely call .select on a control (e.g., tabs). Terminates on error.
Sub SafeSelect(objectPath)
    On Error Resume Next
    Dim obj
    Set obj = SafeFindById(objectPath)
    obj.select
    If Err.Number <> 0 Then
        Dim errMsg
        errMsg = "CRITICAL ERROR: Unable to select '" & objectPath & "'" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
            "This may indicate the SAP GUI is not responding or the session ended. Terminating script..."
        CleanupAndTerminate errMsg
    End If
    On Error GoTo 0
End Sub

' Reads a value by technical field id only.
' Aborts the script with a clear message if the field cannot be read (layout mismatch or restriction).
Function GetCellValueStrict(grid, rowIndex, techId)
    Dim v
    On Error Resume Next
    v = grid.GetCellValue(rowIndex, techId)
    If Err.Number <> 0 Then
        Dim errNum, errDesc
        errNum = Err.Number
        errDesc = Err.Description
        Err.Clear
        CleanupAndTerminate "GetCellValue failed for row " & rowIndex & " on field '" & techId & "'. " & _
            "Layout may differ or field is not readable. Aborting." & vbCrLf & _
            "VBScript Error " & errNum & ": " & errDesc
    End If
    On Error GoTo 0
    GetCellValueStrict = v
End Function

' Stops the script gracefully, performing cleanup and logging the last message to the user.
Sub CleanupAndTerminate(LastMessageToUser)
    ' === CLEANUP ===
    'workbook.Close False 'otherwise might fail at workbook.Close False e.g. when the file was not found
    'excelApp.Quit
    
    Set sheet1 = Nothing
    Set sheet2 = Nothing
    Set workbook = Nothing
    Set excelApp = Nothing
    
    Log LastMessageToUser
    logFile.Close
    WScript.Quit
End Sub
