' Daniel Hermes, started 10. February 2026
' Mechanics & electricians have to log their work in a digital shift log & in SAP in Manchster. It is double work.
' This script parses through an export of the logged work and logs the work in SAP (either fully automatic or asks user to confirm every single SAP interaction, depending on parameter)
Option Explicit ' forces to declare all variables with Dim, Private, or Public
Dim g_logFilePath, logFile, filePath, excelApp, workbook, sheet1, lastRow, lastCol, loadedFromMainScript, fso, file, code, session, autoConfirmResponse, argFilePath, argUseCurrentExcel, argAutoConfirm
Dim sheet1_cached, employeeMapping_cache, g_timezoneBias, g_statusBuffer(), prevScreenUpdating, prevCalculation, column_in_excel_where_to_put_message, done_text_from_excel, cancelled_text_from_excel, SAP_plantcode
initialize()

' DEBUGGING SETTINGS
rowsToInspect = Array() 'PUT HERE WHICH ROWS YOU WANT TO CHECK. example: rowsToInspect = Array(6, 9, 69)
onlyParse_rowsToInspect = False 'TRUE IF YOU ONLY WANT THE SCRIPT TO PARSE THROUGH rowsToInspect. FALSE IF YOU WANT THE SCRIPT TO PARSE THROUGH ALL ROWS BUT LOG SPECIFIC STUFF FOR rowsToInspect

' files:
' https://cargillonline.sharepoint.com/sites/MaintenanceShiftLog/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FMaintenanceShiftLog%2FShared%20Documents%2FGeneral%2FYesterdays%20Log


' === LOOP THROUGH ROWS ===
Dim i, j, SkipReason, rowText, Shift_Start_Date, WO_Nr, Mitarbeiter, Massnahme, DauerInH, Status, emptyRowCounter, rowsToInspect, onlyParse_rowsToInspect
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
        Shift_Start_Date = GetCellFromCache(sheet1_cached, i, 3)
        Mitarbeiter = GetCellFromCache(sheet1_cached, i, 5)
        WO_Nr = GetCellFromCache(sheet1_cached, i, 6)
        Massnahme = GetCellFromCache(sheet1_cached, i, 10)
        DauerInH = GetCellFromCache(sheet1_cached, i, 11)
        Status = GetCellFromCache(sheet1_cached, i, 12)

        ' HARD skip conditions - no WO, already something in message column from a previous script execution or no status/cancelled
        If WO_Nr = "" Then
            SkipReason = SkipReason & vbCrLf & " - No WO found."
        ElseIf Len(WO_Nr) <> 9 Then
            SkipReason = SkipReason & vbCrLf & " - WO " & WO_Nr & " is " & Len(WO_Nr) & " characters long (but should be 9 characters long)."
        End If
        
        'SOFT skip conditions - missing data that is needed for confirmation
        If Mitarbeiter = "" Then
            SkipReason = SkipReason & vbCrLf & " - No employee for WO " & WO_Nr & " found."
        End If
        If Shift_Start_Date = "" Then
            SkipReason = SkipReason & vbCrLf & " - No Shift_Start_Date for WO " & WO_Nr & " found."
        End If
        If DauerInH = "" Then
            SkipReason = SkipReason & vbCrLf & " - No 'DauerInH' for WO " & WO_Nr & " found."
        End If
        If Status = "" Then
            SkipReason = SkipReason & vbCrLf & " - No 'Status' for WO " & WO_Nr & " found."
        End If
        
        If IsInArray(i, rowsToInspect) Then
            Log vbCrLf & _
                "______________________________________________________________________________" & vbCrLf & _
                "Reached row " & i & " to inspect..." & vbCrLf & _
                "Row content:" & vbCrLf & _
                rowText & vbCrLf & _
                "SkipReason so far: " & SkipReason
            'CleanupAndTerminate "Terminating due to reaching a row to inspect (debugging)..."
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
                Dim confirmation
                confirmation = Confirm_WO(WO_Nr, Mitarbeiter, DauerInH, Shift_Start_Date, Massnahme, Status = "True")
                WriteToExcel i, column_in_excel_where_to_put_message, confirmation, False
                Log "confirmation for WO " & WO_Nr & ": " & confirmation
                'Else
                '    Log "WO " & WO_Nr & " does not need confirmation."
            End If
        Else
            Log "Skipping row " & i & " for the following reason(s): " & SkipReason
        End If
    End If
Next

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
        filePath = "https://cargillonline.sharepoint.com/sites/MaintenanceShiftLog/Shared%20Documents/General/Yesterdays%20Log/Yesterdays_log.csv"
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
    Log "Shift logbook Excel file has " & workbook.sheets.Count & " sheets. sheet1 name: " & sheet1.Name
    
    
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
    On Error GoTo 0

    ' --- Load employee -> SAP personnel number mapping from standalone spreadsheet into memory once ---
    Dim mappingFilePath, mappingWb, mappingSheet, mappingExcel, lastRow2, lastCol2
    mappingFilePath = "https://cargillonline.sharepoint.com/sites/MaintenanceShiftLog/Shared%20Documents/General/Yesterdays%20Log/Employee_SAP_mapping.xlsx"
    
    ' Attempt to open the mapping file in a separate, invisible Excel instance so nothing is shown to the user
    On Error Resume Next
    Set mappingExcel = CreateObject("Excel.Application")
    If Not mappingExcel Is Nothing Then
        mappingExcel.Visible = False
        mappingExcel.DisplayAlerts = False
        Set mappingWb = mappingExcel.Workbooks.Open(mappingFilePath)
        If Err.Number <> 0 Or mappingWb Is Nothing Then
            Err.Clear
            Log "Warning: Could not open employee mapping workbook in invisible Excel: " & mappingFilePath & " - lookups will not work."
            If Not mappingExcel Is Nothing Then
                On Error Resume Next
                mappingExcel.Quit
                Set mappingExcel = Nothing
                On Error GoTo 0
            End If
        Else
            Log "Opened employee mapping workbook invisibly. Caching and closing it immediately."
            Set mappingSheet = mappingWb.Sheets(1)
            lastRow2 = mappingSheet.Cells(mappingSheet.Rows.Count, 1).End( - 4162).Row ' xlUp
            lastCol2 = mappingSheet.Cells(1, mappingSheet.Columns.Count).End( - 4159).Column ' xlToLeft
            employeeMapping_cache = Empty
            employeeMapping_cache = mappingSheet.Range(mappingSheet.Cells(1, 1), mappingSheet.Cells(lastRow2, lastCol2)).Value
            If Err.Number <> 0 Then
                Err.Clear
                Log "Warning: Failed to cache employee mapping sheet. Lookups will not work."
                employeeMapping_cache = Empty
            Else
                Log "Cached employee mapping sheet with " & CStr(lastRow2 - 1) & " entries."
            End If
            
            ' Close and quit the invisible Excel instance
            On Error Resume Next
            mappingWb.Close False
            mappingExcel.Quit
            Set mappingWb = Nothing
            Set mappingSheet = Nothing
            Set mappingExcel = Nothing
            On Error GoTo 0
        End If
    Else
        Log "Warning: Could not create an invisible Excel instance to read mapping file. Lookups will not work."
    End If
    On Error GoTo 0
    
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

Function Confirm_WO(wo, mitarbeiter, dauerInStunden, shift_Start_Date, massnahme, finalConfirmation)
    Log "Confirming WO      : " & wo & vbCrLf & _
        " Employee          : " & mitarbeiter & vbCrLf & _
        " Duration (hours)  : " & dauerInStunden & vbCrLf & _
        " Start (UTC)       : " & shift_Start_Date & vbCrLf & _
        " What was done     : " & massnahme & vbCrLf & _
        " finalConfirmation : " & finalConfirmation
    
    
    Dim start_dt, finish_dt
    start_dt = ParseIso8601Z(shift_Start_Date)

    ' B) Start variables
    Dim work_start_date, work_start_time
    work_start_date = FormatDateDDMMYYYY(start_dt) ' => "09.02.2026"
    work_start_time = FormatTimeHHMMSS(start_dt) ' => "23:50:00"

    ' C) Finish variables (add duration in seconds)
    Dim dur_seconds
    dur_seconds = ParseDurationHoursToSeconds(dauerInStunden)
    finish_dt = DateAdd("s", dur_seconds, start_dt)

    Dim work_finish_date, work_finish_time
    work_finish_date = FormatDateDDMMYYYY(finish_dt) ' => "10.02.2026"
    work_finish_time = FormatTimeHHMMSS(finish_dt) ' => "01:20:00"

    ' Demo output (use cscript.exe to see console output)
    Log "work_start_date  = " & work_start_date
    Log "work_start_time  = " & work_start_time
    Log "work_finish_date = " & work_finish_date
    Log "work_finish_time = " & work_finish_time
    
    SafeStartTransaction "IW41"
    ' enable Parameters > Goods movements > all components - see https://cargillonline.sharepoint.com/:i:/r/sites/SAPSUBERLIN/Shared%20Documents/Allgemeines/Anleitungen/Anwendung%20-%20Maintenance%20%26%20Reliability/Korrektive%20Instandhaltung%20(Currative%20Maintenance)/6_Arbeitszeitbest%C3%A4tigung%20(Time%20conformation)/Materialaustrag_Parameter-Einstellung.png
    SafeSendVKey "wnd[0]", 18 ' opens Parameters
    SafeSetSelected "wnd[1]/usr/chkTCORU-ACOMP", True ' ticks Goods movements > all components
    SafeSendVKey "wnd[1]", 0 ' enter
    
    SafeSetText "wnd[0]/usr/ctxtCORUF-AUFNR", wo
    Log "Assuming work relates to the first operation (0010)."
    SafeFindById("wnd[0]/usr/txtCORUF-VORNR").text = "0010"
    SafeSendVKey "wnd[0]", 0 'enter
    
    
    Dim personnelNo, duration
    On Error Resume Next
    personnelNo = GetpersonnelNumber(mitarbeiter)
    If Err.Number <> 0 Then
        Confirm_WO = "ERROR: " & Err.Description & vbCrLf & "Personell no. field not found."
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
    SafeSetText "wnd[0]/usr/txtAFRUD-ISMNW_2", dauerInStunden ' Actual work
    SafeSetText "wnd[0]/usr/txtAFRUD-IDAUR", TrimAfterComma(dauerInStunden) ' Actual duration - Actual work can be X,YZ while Actual duration can only be one char after comma e.g. X,Y
    SafeSetText "wnd[0]/usr/ctxtAFRUD-ISDD", work_start_date ' work start day
    SafeSetText "wnd[0]/usr/ctxtAFRUD-ISDZ", work_start_time ' work start time
    SafeSetText "wnd[0]/usr/ctxtAFRUD-IEDD", work_finish_date ' work end day
    SafeSetText "wnd[0]/usr/ctxtAFRUD-IEDZ", work_finish_time ' work end time
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
        SafeSetText "wnd[0]/usr/txtAFRUD-OFMNW_2", "0" ' Set "Remaining work" field to 0
        SafeSetSelected "wnd[0]/usr/chkAFRUD-AUERU", True ' Tick "Final confirmation" checkbox
        SafeSetSelected "wnd[0]/usr/chkAFRUD-LEKNW", True ' Tick "No remaining work" checkbox
    End If
    
    WScript.Quit
    
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
    ' Search in cached employee mapping for matching employee name and return SAP personnel number
    Dim r, employeeName_in_cache, personnelNo
    
    If Not IsEmpty(employeeMapping_cache) And Not IsNull(employeeMapping_cache) Then
        ' Iterate through mapping rows (starting from row 2, assuming row 1 is header)
        For r = 2 To UBound(employeeMapping_cache, 1)
            employeeName_in_cache = GetCellFromCache(employeeMapping_cache, r, 1) ' Get name from column 1

            ' Compare normalized names - allow lenient matching
            'If CompareStringsLeniently(employeeName_Normalized, employeeName_in_cache_Normalized) Then
            
            If StrComp(employeeName, employeeName_in_cache, vbTextCompare) = 0 Then ' case-insensitive exact match
                personnelNo = GetCellFromCache(employeeMapping_cache, r, 2) ' Get SAP personnel number from column 2 (=column B)
                If personnelNo <> "" Then
                    Log "Employee '" & employeeName & "' found in mapping, row " & r & " as '" & employeeName_in_cache & "', mapped to personnel number: " & personnelNo
                    GetpersonnelNumber = personnelNo
                    Exit Function
                Else
                    CleanupAndTerminate "Employee '" & employeeName & "' found in mapping, row " & r & " as '" & employeeName_in_cache & "', but personnel number field is empty."
                End If
            End If
        Next
    Else Log "Employee mapping cache is empty. Cannot perform employee name to personnel number lookups." End If

    ' If no match found in mapping, log and return empty (don't confirm with personnel number)
    Log vbCrLf & "=== WARNING === Unknown employee: " & employeeName & " // normalized to: " & employeeName_Normalized & " // Not found in mapping. Confirming without entering employee to SAP..."
End Function

'=== Strictly parse ISO 8601 UTC: "YYYY-MM-DDTHH:MM:SSZ" ===
Function ParseIso8601Z(s)
    Dim re, m
    Set re = New RegExp
    re.Pattern = "^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})Z$"
    re.Global = False
    re.IgnoreCase = False

    If Not re.Test(s) Then
        Err.Raise vbObjectError + 1, , "start_datetime must be exactly 'YYYY-MM-DDTHH:MM:SSZ'. Got: " & s
    End If

    Set m = re.Execute(s)(0)

    Dim yr, mo, dy, hh, nn, ss
    yr = CInt(m.SubMatches(0))
    mo = CInt(m.SubMatches(1))
    dy = CInt(m.SubMatches(2))
    hh = CInt(m.SubMatches(3))
    nn = CInt(m.SubMatches(4))
    ss = CInt(m.SubMatches(5))

    ' Range checks
    If mo < 1 Or mo > 12 Then Err.Raise vbObjectError + 2, , "Invalid month: " & mo
    If dy < 1 Or dy > 31 Then Err.Raise vbObjectError + 3, , "Invalid day: " & dy
    If hh < 0 Or hh > 23 Then Err.Raise vbObjectError + 4, , "Invalid hour: " & hh
    If nn < 0 Or nn > 59 Then Err.Raise vbObjectError + 5, , "Invalid minute: " & nn
    If ss < 0 Or ss > 59 Then Err.Raise vbObjectError + 6, , "Invalid second: " & ss

    Dim d, t, dt
    d = DateSerial(yr, mo, dy)
    If Year(d) <> yr Or Month(d) <> mo Or Day(d) <> dy Then
        Err.Raise vbObjectError + 7, , "Invalid calendar date in input: " & s
    End If

    t = TimeSerial(hh, nn, ss)
    dt = d + t
    ParseIso8601Z = dt ' VBScript Date (interpreted as UTC conceptually)
End Function

'=== Utility: zero-pad ===
Function Pad2(n)
    Pad2 = Right("0" & CStr(n), 2)
End Function

'=== Formatters ===
Function FormatDateDDMMYYYY(d)
    FormatDateDDMMYYYY = Pad2(Day(d)) & "." & Pad2(Month(d)) & "." & Year(d)
End Function

Function FormatTimeHHMMSS(d)
    FormatTimeHHMMSS = Pad2(Hour(d)) & ":" & Pad2(Minute(d)) & ":" & Pad2(Second(d))
End Function

'=== Parse duration (hours as a number string) into seconds ===
' Accepts "3", "1.5", "0", "2.25", etc. Uses dot as decimal separator.
Function ParseDurationHoursToSeconds(s)
    Dim trimmed, hoursDbl, seconds
    trimmed = Trim(s)

    ' Optional: normalize comma decimals (uncomment next line if you want to allow "1,5")
    ' trimmed = Replace(trimmed, ",", ".")

    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 20, , "work_duration cannot be empty"
    End If

    If Not IsNumeric(trimmed) Then
        Err.Raise vbObjectError + 21, , "work_duration must be a number of hours (e.g., 1.5 or 3). Got: " & s
    End If

    hoursDbl = CDbl(trimmed)
    If hoursDbl < 0 Then
        Err.Raise vbObjectError + 22, , "work_duration cannot be negative. Got: " & s
    End If

    ' Convert hours -> seconds, rounding to nearest second to avoid FP issues
    seconds = CLng(hoursDbl * 3600 + 0.5)
    ParseDurationHoursToSeconds = seconds
End Function

Function DeleteNewLineAndTrim(inputString)
    inputString = Replace(inputString, vbCrLf, "")
    inputString = Replace(inputString, vbCr, "")
    inputString = Replace(inputString, vbLf, "")
    inputString = Trim(inputString) ' ignore leading & lagging spaces
    DeleteNewLineAndTrim = inputString
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