Attribute VB_Name = "AML_Period_Comparison"
' =========================================================================
' [AML] TRANSACTION MONITORING: LOOKBACK PERIOD BATCH RECONCILIATION COCKPIT
' =========================================================================
' Version: 5.0 (Restored Classic Layout, Pixel-Perfect Alignment & Export)
'
' Key Capabilities:
'  1. RESTORED CLASSIC DASHBOARD: Fully restores the beloved 14-column layout
'     from Version 3.0 (including % Change, ECM ID, Alert ID, and all 5 KPI cards:
'     Evaluated, Increased Count, Total Txns, New Alerted, Alerts Unchanged).
'  2. PIXEL-PERFECT ALIGNMENT: 5 balanced action buttons on Row 12 directly align
'     with the 5 KPI tiles above and the 14-column table below.
'  3. PIVOT COUNTERPARTY DIFFING: Checks Pivot tables between Old and New files
'     to identify newly added counterparties (e.g. MONICA HERZBERG CLAUDIO EULAU).
'  4. ELEGANT & ALIGNED DISCREPANCY AUDIT: Every alert section in Discrepancy_Added_Txns
'     spans the exact same columns (B to I) with fixed widths, Segoe UI typography,
'     crisp borders, and a dedicated 'Pivot Status' column.
'  5. ONE-CLICK EXPORT TO WORKBOOK: Interactive button exports a clean, standalone,
'     formatted .xlsx workbook containing the Dashboard and Discrepancy details.
'  6. 100% NATIVE EXCEL VBA: Zero ActiveX/Object error 438, zero repair prompts.
' =========================================================================
Option Explicit

' --- Design System Palette (Slate Theme) ---
Private Const COLOR_HEADER_BG As Long = 2762511     ' Slate 900 RGB(15, 23, 42)
Private Const COLOR_HEADER_TXT As Long = 16777215   ' White     RGB(255, 255, 255)
Private Const COLOR_BORDER As Long = 15790320       ' Slate 200 RGB(226, 232, 240)
Private Const COLOR_CARD_BG As Long = 16777215      ' White     RGB(255, 255, 255)
Private Const COLOR_ACCENT_BLUE As Long = 14399778  ' Sky 600   RGB(37, 99, 235)
Private Const COLOR_ALERT_RED As Long = 4079854     ' Red 600   RGB(220, 38, 38)
Private Const COLOR_SUCCESS_GREEN As Long = 3256086 ' Green 600 RGB(22, 163, 74)
Private Const COLOR_AMBER As Long = 2727916         ' Amber 600 RGB(217, 119, 6)

' --- Dashboard Sheet Names ---
Private Const SHEET_DASHBOARD As String = "Comparison Dashboard"
Private Const SHEET_ADDED_TXNS As String = "Discrepancy_Added_Txns"
Private Const SHEET_DROPPED_TXNS As String = "Discrepancy_Dropped_Txns"

' =========================================================================
' [UI] 0. USER-FACING MACRO: CREATE / SETUP THE DASHBOARD
' =========================================================================
Public Sub Setup_Dashboard()
    Dim ws As Worksheet
    Set ws = Setup_Comparison_Dashboard(ActiveWorkbook, True)
    ws.Activate
    MsgBox "Comparison Dashboard has been created successfully!" & vbCrLf & vbCrLf & _
           "Click the interactive buttons on the sheet to run batch comparisons or export results.", vbInformation, "Dashboard Ready"
End Sub

' =========================================================================
' [RUN] 1. USER-FACING MACRO: RUN BATCH COMPARISON ACROSS ALL ALERTS
' =========================================================================
Public Sub Run_Batch_Comparison()
    Dim selectedFolder As String, oldFolder As String, newFolder As String
    Dim pSep As String
    Dim wbDashboard As Workbook
    Dim wsDash As Worksheet
    
    Set wbDashboard = ActiveWorkbook
    Set wsDash = Setup_Comparison_Dashboard(wbDashboard)
    pSep = Application.PathSeparator
    
    ' Smart Auto-Detection: check if active workbook directory contains Old/ and New/
    If wbDashboard.Path <> "" Then
        If (Dir(wbDashboard.Path & pSep & "Old", vbDirectory) <> "" Or Dir(wbDashboard.Path & pSep & "old", vbDirectory) <> "") And _
           (Dir(wbDashboard.Path & pSep & "New", vbDirectory) <> "" Or Dir(wbDashboard.Path & pSep & "new", vbDirectory) <> "") Then
            oldFolder = wbDashboard.Path & pSep & "Old"
            If Dir(oldFolder, vbDirectory) = "" Then oldFolder = wbDashboard.Path & pSep & "old"
            newFolder = wbDashboard.Path & pSep & "New"
            If Dir(newFolder, vbDirectory) = "" Then newFolder = wbDashboard.Path & pSep & "new"
            Execute_Batch_Comparison wbDashboard, oldFolder, newFolder
            Exit Sub
        End If
    End If
    
    ' Step 1: Prompt for parent folder or Old folder
    selectedFolder = Pick_Folder("Select Root Folder (containing 'Old' & 'New' subfolders) OR select the 'Old' folder:")
    If selectedFolder = "" Then Exit Sub
    
    If Dir(selectedFolder & pSep & "Old", vbDirectory) <> "" And Dir(selectedFolder & pSep & "New", vbDirectory) <> "" Then
        oldFolder = selectedFolder & pSep & "Old"
        newFolder = selectedFolder & pSep & "New"
    ElseIf Dir(selectedFolder & pSep & "old", vbDirectory) <> "" And Dir(selectedFolder & pSep & "new", vbDirectory) <> "" Then
        oldFolder = selectedFolder & pSep & "old"
        newFolder = selectedFolder & pSep & "new"
    Else
        oldFolder = selectedFolder
        newFolder = Pick_Folder("Now select the 'New' lookback period folder:")
        If newFolder = "" Then Exit Sub
    End If
    
    Execute_Batch_Comparison wbDashboard, oldFolder, newFolder
End Sub

' =========================================================================
' [RUN] 2. USER-FACING MACRO: RUN SINGLE PAIR COMPARISON
' =========================================================================
Public Sub Run_Single_Pair_Comparison()
    Dim oldFile As String, newFile As String
    Dim wbDashboard As Workbook
    
    Set wbDashboard = ActiveWorkbook
    Setup_Comparison_Dashboard wbDashboard
    
    oldFile = Pick_Excel_File("Select the OLD / In-Scope Lookback Transaction File:")
    If oldFile = "" Then Exit Sub
    
    newFile = Pick_Excel_File("Select the NEW / Expanded Lookback Transaction File:")
    If newFile = "" Then Exit Sub
    
    Execute_Single_Comparison wbDashboard, oldFile, newFile
End Sub

' =========================================================================
' [EXPORT] 3. USER-FACING MACRO: EXPORT TO STANDALONE FORMATTED WORKBOOK
' =========================================================================
Public Sub Export_To_New_Workbook()
    Dim wbSource As Workbook, wbNew As Workbook
    Dim wsDash As Worksheet, wsAdded As Worksheet, wsDropped As Worksheet
    Dim defaultFileName As String, savePath As Variant
    Dim timeStamp As String, shp As Shape
    
    Set wbSource = ActiveWorkbook
    On Error Resume Next
    Set wsDash = wbSource.Worksheets(SHEET_DASHBOARD)
    On Error GoTo 0
    
    If wsDash Is Nothing Then
        MsgBox "Comparison Dashboard sheet not found. Please run a comparison first!", vbExclamation, "No Data to Export"
        Exit Sub
    End If
    
    timeStamp = Format(Now, "yyyymmdd_hhnnss")
    defaultFileName = "AML_Lookback_Comparison_Report_" & timeStamp & ".xlsx"
    
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, _
                                            FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
                                            Title:="Export AML Comparison Report to New Workbook")
                                            
    If VarType(savePath) = vbBoolean And savePath = False Then Exit Sub ' User cancelled
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Copy Dashboard to new workbook
    wsDash.Copy
    Set wbNew = ActiveWorkbook
    
    ' Strip action button shapes from exported dashboard for clean executive view
    On Error Resume Next
    For Each shp In wbNew.Worksheets(1).Shapes
        If Left(shp.Name, 4) = "Btn_" Then shp.Delete
    Next shp
    On Error GoTo 0
    
    ' Copy Discrepancy Added Txns
    On Error Resume Next
    Set wsAdded = wbSource.Worksheets(SHEET_ADDED_TXNS)
    On Error GoTo 0
    If Not wsAdded Is Nothing Then
        wsAdded.Copy After:=wbNew.Worksheets(wbNew.Worksheets.Count)
    End If
    
    ' Copy Discrepancy Dropped Txns (if exists)
    On Error Resume Next
    Set wsDropped = wbSource.Worksheets(SHEET_DROPPED_TXNS)
    On Error GoTo 0
    If Not wsDropped Is Nothing Then
        wsDropped.Copy After:=wbNew.Worksheets(wbNew.Worksheets.Count)
    End If
    
    wbNew.Worksheets(1).Activate
    wbNew.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "AML Lookback Reconciliation Report successfully exported to:" & vbCrLf & vbCrLf & _
           CStr(savePath), vbInformation, "Export Successful"
End Sub

' =========================================================================
' [RESET] 4. USER-FACING MACRO: RESET DASHBOARD
' =========================================================================
Public Sub Clear_Comparison_Dashboard()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook
    Set ws = Setup_Comparison_Dashboard(wb, True)
    
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(SHEET_ADDED_TXNS).Delete
    wb.Worksheets(SHEET_DROPPED_TXNS).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    MsgBox "Dashboard has been reset.", vbInformation, "Cleared"
End Sub

' =========================================================================
' [VIEW] 5. USER-FACING MACRO: VIEW GRANULAR DISCREPANCIES
' =========================================================================
Public Sub View_Added_Transactions()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    On Error Resume Next
    wb.Worksheets(SHEET_ADDED_TXNS).Activate
    If Err.Number <> 0 Then
        MsgBox "No added transactions sheet exists yet. Please run a comparison first!", vbExclamation, "No Audit Data"
    End If
    On Error GoTo 0
End Sub

' =========================================================================
' [ENGINE] 6. CORE BATCH COMPARISON ENGINE
' =========================================================================
Private Sub Execute_Batch_Comparison(ByVal wbDash As Workbook, ByVal oldDir As String, ByVal newDir As String)
    Dim wsDash As Worksheet
    Dim wsAdded As Worksheet, wsDropped As Worksheet
    Dim oldFiles As Collection, newFiles As Collection
    Dim oldMap As Object, newMap As Object
    Dim alertKeys As Collection
    Dim curRow As Long
    Dim alertKey As String, ecmID As String, alertID As String
    Dim oldFilePath As String, newFilePath As String
    Dim totalAlerts As Long, countIncreasedAlerts As Long, countUnchangedAlerts As Long
    Dim totalAddedTxns As Long, totalDroppedTxns As Long, totalAlertedAdded As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wsDash = wbDash.Worksheets(SHEET_DASHBOARD)
    Set wsAdded = GetOrCreateWorksheet(wbDash, SHEET_ADDED_TXNS)
    Set wsDropped = GetOrCreateWorksheet(wbDash, SHEET_DROPPED_TXNS)
    
    FormatAuditSheetHeaders wsAdded, "NEWLY ADDED TRANSACTIONS GROUPED BY ALERT (EXPANDED LOOKBACK PERIOD)", True
    FormatAuditSheetHeaders wsDropped, "DROPPED / EXCLUDED TRANSACTIONS GROUPED BY ALERT", False
    
    Set oldFiles = ListExcelFiles(oldDir)
    Set newFiles = ListExcelFiles(newDir)
    
    If oldFiles.Count = 0 Then
        MsgBox "No Excel files (.xlsx / .xlsm) found in Old Folder: " & vbCrLf & oldDir, vbExclamation, "No Files Found"
        GoTo CleanExit
    End If
    
    If newFiles.Count = 0 Then
        MsgBox "No Excel files (.xlsx / .xlsm) found in New Folder: " & vbCrLf & newDir, vbExclamation, "No Files Found"
        GoTo CleanExit
    End If
    
    Set oldMap = CreateLookupDict()
    Set newMap = CreateLookupDict()
    Set alertKeys = New Collection
    
    Dim fPath As Variant
    For Each fPath In oldFiles
        alertKey = ExtractMatchKey(CStr(fPath))
        If Not DictExists(oldMap, alertKey) Then
            DictAdd oldMap, alertKey, CStr(fPath)
            alertKeys.Add alertKey
        End If
    Next fPath
    
    For Each fPath In newFiles
        alertKey = ExtractMatchKey(CStr(fPath))
        If Not DictExists(newMap, alertKey) Then
            DictAdd newMap, alertKey, CStr(fPath)
            If Not DictExists(oldMap, alertKey) Then alertKeys.Add alertKey
        End If
    Next fPath
    
    curRow = 15
    wsDash.Range("B15:O500").ClearContents
    wsDash.Range("B15:O500").ClearFormats
    
    Dim keyVar As Variant
    For Each keyVar In alertKeys
        alertKey = CStr(keyVar)
        oldFilePath = ""
        newFilePath = ""
        If DictExists(oldMap, alertKey) Then oldFilePath = CStr(DictGet(oldMap, alertKey))
        If DictExists(newMap, alertKey) Then newFilePath = CStr(DictGet(newMap, alertKey))
        
        If oldFilePath <> "" And newFilePath <> "" Then
            totalAlerts = totalAlerts + 1
            ecmID = ExtractECMID(newFilePath)
            If ecmID = "-" Then ecmID = ExtractECMID(oldFilePath)
            alertID = ExtractAlertID(newFilePath)
            If alertID = "" Then alertID = ExtractAlertID(oldFilePath)
            
            ProcessPair wsDash, wsAdded, wsDropped, curRow, ecmID, alertID, oldFilePath, newFilePath, _
                        countIncreasedAlerts, countUnchangedAlerts, totalAddedTxns, totalDroppedTxns, totalAlertedAdded
            curRow = curRow + 1
            
        ElseIf oldFilePath <> "" And newFilePath = "" Then
            totalAlerts = totalAlerts + 1
            ecmID = ExtractECMID(oldFilePath)
            alertID = ExtractAlertID(oldFilePath)
            WriteUnmatchedRow wsDash, curRow, ecmID, alertID, GetFileName(oldFilePath), "(Missing in New Folder)", "MISSING IN NEW [WARN]"
            curRow = curRow + 1
            
        ElseIf oldFilePath = "" And newFilePath <> "" Then
            totalAlerts = totalAlerts + 1
            ecmID = ExtractECMID(newFilePath)
            alertID = ExtractAlertID(newFilePath)
            WriteUnmatchedRow wsDash, curRow, ecmID, alertID, "(Missing in Old Folder)", GetFileName(newFilePath), "NEW UNPAIRED ALERT [WARN]"
            curRow = curRow + 1
        End If
    Next keyVar
    
    FormatBatchTable wsDash, 15, curRow - 1
    UpdateKPICards wsDash, totalAlerts, countIncreasedAlerts, totalAddedTxns, totalAlertedAdded, countUnchangedAlerts
    
    ApplyAuditSheetColumnWidths wsAdded, True
    ApplyAuditSheetColumnWidths wsDropped, False
    wsDash.Activate
    
    MsgBox "Batch Comparison Complete!" & vbCrLf & vbCrLf & _
           "- Total Alerts Evaluated: " & totalAlerts & vbCrLf & _
           "- Alerts With Increased Count: " & countIncreasedAlerts & vbCrLf & _
           "- Alerts With No Count Change: " & countUnchangedAlerts & vbCrLf & _
           "- Total Transactions Added: " & totalAddedTxns & vbCrLf & _
           "- New Alerted Transactions: " & totalAlertedAdded, vbInformation, "Reconciliation Finished"

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error during Batch Comparison: " & Err.Description, vbCritical, "Error " & Err.Number
    Resume CleanExit
End Sub

' =========================================================================
' [ENGINE] 7. SINGLE PAIR COMPARISON EXECUTION
' =========================================================================
Private Sub Execute_Single_Comparison(ByVal wbDash As Workbook, ByVal oldFilePath As String, ByVal newFilePath As String)
    Dim wsDash As Worksheet
    Dim wsAdded As Worksheet, wsDropped As Worksheet
    Dim ecmID As String, alertID As String
    Dim countIncreasedAlerts As Long, countUnchangedAlerts As Long
    Dim totalAddedTxns As Long, totalDroppedTxns As Long, totalAlertedAdded As Long
    Dim curRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wsDash = wbDash.Worksheets(SHEET_DASHBOARD)
    Set wsAdded = GetOrCreateWorksheet(wbDash, SHEET_ADDED_TXNS)
    Set wsDropped = GetOrCreateWorksheet(wbDash, SHEET_DROPPED_TXNS)
    
    FormatAuditSheetHeaders wsAdded, "NEWLY ADDED TRANSACTIONS GROUPED BY ALERT (EXPANDED LOOKBACK PERIOD)", True
    FormatAuditSheetHeaders wsDropped, "DROPPED / EXCLUDED TRANSACTIONS GROUPED BY ALERT", False
    
    ecmID = ExtractECMID(newFilePath)
    alertID = ExtractAlertID(newFilePath)
    curRow = 15
    
    wsDash.Range("B15:O30").ClearContents
    wsDash.Range("B15:O30").ClearFormats
    
    ProcessPair wsDash, wsAdded, wsDropped, curRow, ecmID, alertID, oldFilePath, newFilePath, _
                countIncreasedAlerts, countUnchangedAlerts, totalAddedTxns, totalDroppedTxns, totalAlertedAdded
                
    FormatBatchTable wsDash, 15, 15
    UpdateKPICards wsDash, 1, countIncreasedAlerts, totalAddedTxns, totalAlertedAdded, countUnchangedAlerts
    
    ApplyAuditSheetColumnWidths wsAdded, True
    ApplyAuditSheetColumnWidths wsDropped, False
    wsDash.Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Single Alert Comparison Complete!" & vbCrLf & vbCrLf & _
           "ECM ID: " & ecmID & " | Alert ID: " & alertID & vbCrLf & _
           "- Net Count Delta: " & IIf(totalAddedTxns - totalDroppedTxns >= 0, "+", "") & (totalAddedTxns - totalDroppedTxns) & vbCrLf & _
           "- Newly Added Alerted Transactions: " & totalAlertedAdded, vbInformation, "Reconciliation Finished"

CleanExit:
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error during Single Comparison: " & Err.Description, vbCritical, "Error " & Err.Number
End Sub

' =========================================================================
' [PROCESS] 8. PROCESS INDIVIDUAL ALERT PAIR
' =========================================================================
Private Sub ProcessPair(ByVal wsDash As Worksheet, ByVal wsAdded As Worksheet, ByVal wsDropped As Worksheet, _
                        ByVal rowIdx As Long, ByVal ecmID As String, ByVal alertID As String, _
                        ByVal oldFilePath As String, ByVal newFilePath As String, _
                        ByRef countIncreasedAlerts As Long, ByRef countUnchangedAlerts As Long, _
                        ByRef totalAddedTxns As Long, ByRef totalDroppedTxns As Long, _
                        ByRef totalAlertedAdded As Long)
                        
    Dim wbOld As Workbook, wbNew As Workbook
    Dim wsOld As Worksheet, wsNew As Worksheet
    Dim oldMap As Object, newMap As Object
    Dim oldPivotCPs As Object, newPivotCPs As Object
    Dim newCPList As New Collection
    Dim oldTxCount As Long, newTxCount As Long, countDelta As Long
    Dim oldAlertCount As Long, newAlertCount As Long, alertDelta As Long
    Dim oldMinDate As Variant, oldMaxDate As Variant
    Dim newMinDate As Variant, newMaxDate As Variant
    Dim oldPeriodStr As String, newPeriodStr As String
    Dim statusStr As String
    Dim pctChange As Double
    
    Set wbOld = Workbooks.Open(oldFilePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsOld = FindTransactionSheet(wbOld)
    Set oldMap = ExtractTransactions(wsOld, oldTxCount, oldAlertCount, oldMinDate, oldMaxDate)
    Set oldPivotCPs = ExtractPivotCounterparties(wbOld)
    
    Set wbNew = Workbooks.Open(newFilePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsNew = FindTransactionSheet(wbNew)
    Set newMap = ExtractTransactions(wsNew, newTxCount, newAlertCount, newMinDate, newMaxDate)
    Set newPivotCPs = ExtractPivotCounterparties(wbNew)
    
    wbOld.Close SaveChanges:=False
    wbNew.Close SaveChanges:=False
    
    ' Identify newly added counterparties from Pivot tables
    Dim cpKeyVar As Variant, cpUpper As String
    For Each cpKeyVar In DictKeys(newPivotCPs)
        cpUpper = CStr(cpKeyVar)
        If Not DictExists(oldPivotCPs, cpUpper) Then
            newCPList.Add DictGet(newPivotCPs, cpUpper)
        End If
    Next cpKeyVar
    
    ' Two-level Date Resolution: If sheet dates are missing or N/A, extract from filename
    oldPeriodStr = FormatDateRange(oldMinDate, oldMaxDate)
    If oldPeriodStr = "N/A" Then oldPeriodStr = ExtractDateFromFilename(GetFileName(oldFilePath))
    If oldPeriodStr = "" Then oldPeriodStr = "N/A"
    
    newPeriodStr = FormatDateRange(newMinDate, newMaxDate)
    If newPeriodStr = "N/A" Then newPeriodStr = ExtractDateFromFilename(GetFileName(newFilePath))
    If newPeriodStr = "" Then newPeriodStr = "N/A"
    
    countDelta = newTxCount - oldTxCount
    alertDelta = newAlertCount - oldAlertCount
    pctChange = 0#
    If oldTxCount > 0 Then pctChange = (countDelta / oldTxCount) * 100#
    
    If countDelta > 0 Then
        countIncreasedAlerts = countIncreasedAlerts + 1
        If alertDelta > 0 Then
            statusStr = "COUNT INCREASED (+" & countDelta & ") | NEW ALERTED (+" & alertDelta & ") [ALERT]"
        ElseIf newCPList.Count > 0 Then
            statusStr = "COUNT INCREASED (+" & countDelta & ") | NEW CP (+" & newCPList.Count & ") [ALERT]"
        Else
            statusStr = "COUNT INCREASED (+" & countDelta & ") [WARN]"
        End If
    ElseIf countDelta < 0 Then
        statusStr = "COUNT DECREASED (" & countDelta & ") [-]"
    Else
        countUnchangedAlerts = countUnchangedAlerts + 1
        If newCPList.Count > 0 Then
            statusStr = "NO COUNT CHANGE (0) | NEW CP (+" & newCPList.Count & ") [ALERT]"
        Else
            statusStr = "NO COUNT CHANGE (0) [OK]"
        End If
    End If
    
    ' Write Dashboard Table Row (Columns B to O - Exactly 14 columns from Version 3.0)
    With wsDash
        .Cells(rowIdx, 2).Value = ecmID
        .Cells(rowIdx, 3).Value = alertID
        .Cells(rowIdx, 4).Value = GetFileName(oldFilePath)
        .Cells(rowIdx, 5).Value = GetFileName(newFilePath)
        .Cells(rowIdx, 6).Value = oldPeriodStr
        .Cells(rowIdx, 7).Value = newPeriodStr
        .Cells(rowIdx, 8).Value = oldTxCount
        .Cells(rowIdx, 9).Value = newTxCount
        .Cells(rowIdx, 10).Value = countDelta
        .Cells(rowIdx, 11).Value = pctChange / 100#
        .Cells(rowIdx, 12).Value = oldAlertCount
        .Cells(rowIdx, 13).Value = newAlertCount
        .Cells(rowIdx, 14).Value = alertDelta
        .Cells(rowIdx, 15).Value = statusStr
    End With
    
    ' Record newly added transactions grouped by alert (with Pivot diff integrated)
    Dim pairAddedCount As Long, pairAlertedAdded As Long
    RecordAddedTransactions wsAdded, ecmID, alertID, newMap, oldMap, newCPList, pairAddedCount, pairAlertedAdded
    
    Dim pairDroppedCount As Long
    RecordDroppedTransactions wsDropped, ecmID, alertID, oldMap, newMap, pairDroppedCount
    
    totalAddedTxns = totalAddedTxns + pairAddedCount
    totalDroppedTxns = totalDroppedTxns + pairDroppedCount
    totalAlertedAdded = totalAlertedAdded + pairAlertedAdded
End Sub

' =========================================================================
' [PIVOT] 9. EXTRACT COUNTERPARTIES FROM PIVOT TABLE
' =========================================================================
Private Function ExtractPivotCounterparties(ByVal wb As Workbook) As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim pt As PivotTable, pf As PivotField, pi As PivotItem
    Dim foundViaPivot As Boolean
    
    Set dict = CreateLookupDict()
    
    On Error Resume Next
    Set ws = wb.Worksheets("Pivot")
    If ws Is Nothing Then
        Dim s As Worksheet
        For Each s In wb.Worksheets
            If InStr(1, s.Name, "pivot", vbTextCompare) > 0 Then
                Set ws = s
                Exit For
            End If
        Next s
    End If
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ExtractPivotCounterparties = dict
        Exit Function
    End If
    
    ' Strategy 1: Native PivotTable COM Object
    foundViaPivot = False
    On Error Resume Next
    If ws.PivotTables.Count > 0 Then
        For Each pt In ws.PivotTables
            For Each pf In pt.PivotFields
                If InStr(1, pf.Name, "counterparty", vbTextCompare) > 0 Then
                    For Each pi In pf.PivotItems
                        Dim piName As String
                        piName = Trim(pi.Name)
                        If piName <> "" And piName <> "(blank)" And piName <> "(empty)" Then
                            If Not DictExists(dict, UCase(piName)) Then
                                DictAdd dict, UCase(piName), piName
                                foundViaPivot = True
                            End If
                        End If
                    Next pi
                End If
            Next pf
        Next pt
    End If
    On Error GoTo 0
    
    ' Strategy 2: Resilient Column A Scan (Starting row 4)
    If Not foundViaPivot Or DictCount(dict) = 0 Then
        Dim lastRow As Long, r As Long
        Dim cellVal As String, upperVal As String
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For r = 4 To lastRow
            cellVal = Trim(CStr(ws.Cells(r, 1).Value))
            upperVal = UCase(cellVal)
            If upperVal <> "" And _
               upperVal <> "ROW LABELS" And _
               upperVal <> "GRAND TOTAL" And _
               upperVal <> "TOTAL" And _
               upperVal <> "CR" And _
               upperVal <> "DR" And _
               upperVal <> "CR TOTAL" And _
               upperVal <> "DR TOTAL" And _
               upperVal <> "(BLANK)" And _
               InStr(1, upperVal, "MANY TO ONE", vbTextCompare) = 0 And _
               InStr(1, upperVal, "SCENARIO", vbTextCompare) = 0 Then
               
                If Not DictExists(dict, upperVal) Then
                    DictAdd dict, upperVal, cellVal
                End If
            End If
        Next r
    End If
    
    Set ExtractPivotCounterparties = dict
End Function

' =========================================================================
' [METRICS] 10. EXTRACT TRANSACTIONS FROM SHEET
' =========================================================================
Private Function ExtractTransactions(ByVal ws As Worksheet, ByRef txCount As Long, ByRef alertCount As Long, _
                                     ByRef minDate As Variant, ByRef maxDate As Variant) As Object
    Dim dict As Object
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long, hRow As Long
    Dim txIdCol As Long, isAlertCol As Long, dateCol As Long, amtCol As Long, accCol As Long, descCol As Long, cpCol As Long
    Dim hText As String
    Dim txID As String, isAlert As String, dtVal As Variant, amtVal As Double
    Dim accNo As String, descText As String, cpName As String
    Dim curDate As Date, parsedDate As Variant
    Dim recArray As Variant
    
    Set dict = CreateLookupDict()
    txCount = 0
    alertCount = 0
    minDate = Empty
    maxDate = Empty
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Set ExtractTransactions = dict
        Exit Function
    End If
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    ' Scan rows 1 to 3 to locate header row
    hRow = 1
    txIdCol = 1: isAlertCol = 0: dateCol = 0: amtCol = 0: accCol = 0: descCol = 0: cpCol = 0
    Dim testRow As Long
    For testRow = 1 To 3
        For c = 1 To lastCol
            hText = LCase(Trim(CStr(ws.Cells(testRow, c).Value)))
            If InStr(hText, "transaction id") > 0 Or InStr(hText, "txn id") > 0 Then
                txIdCol = c: hRow = testRow
            ElseIf InStr(hText, "is alerted") > 0 Or hText = "alerted" Then
                isAlertCol = c: hRow = testRow
            ElseIf (InStr(hText, "date") > 0 Or InStr(hText, "dt") > 0) And dateCol = 0 Then
                dateCol = c: hRow = testRow
            ElseIf (InStr(hText, "amount") > 0 Or InStr(hText, "value") > 0) And amtCol = 0 Then
                amtCol = c: hRow = testRow
            ElseIf InStr(hText, "account") > 0 And accCol = 0 Then
                accCol = c: hRow = testRow
            ElseIf (InStr(hText, "description") > 0 Or InStr(hText, "desc") > 0) And descCol = 0 Then
                descCol = c: hRow = testRow
            ElseIf InStr(hText, "counterparty") > 0 And cpCol = 0 Then
                cpCol = c: hRow = testRow
            End If
        Next c
        If dateCol > 0 Then Exit For
    Next testRow
    
    ' Smart Fallback defaults based on AML schema
    If dateCol = 0 Then
        For c = 1 To lastCol
            If InStr(LCase(Trim(CStr(ws.Cells(hRow, c).Value))), "date") > 0 Then
                dateCol = c: Exit For
            End If
        Next c
    End If
    If isAlertCol = 0 Then isAlertCol = 2
    
    For r = hRow + 1 To lastRow
        txID = Trim(CStr(ws.Cells(r, txIdCol).Value))
        If txID <> "" Then
            txCount = txCount + 1
            
            isAlert = "No"
            If isAlertCol > 0 Then
                isAlert = Trim(CStr(ws.Cells(r, isAlertCol).Value))
                If LCase(isAlert) = "yes" Or LCase(isAlert) = "y" Or isAlert = "1" Or LCase(isAlert) = "true" Then
                    alertCount = alertCount + 1
                    isAlert = "Yes"
                Else
                    isAlert = "No"
                End If
            End If
            
            ' Robust Date Parsing
            dtVal = ws.Cells(r, dateCol).Value
            parsedDate = ParseDateValue(dtVal)
            If Not IsEmpty(parsedDate) Then
                curDate = CDate(parsedDate)
                If IsEmpty(minDate) Or curDate < minDate Then minDate = curDate
                If IsEmpty(maxDate) Or curDate > maxDate Then maxDate = curDate
            End If
            
            amtVal = 0#
            If amtCol > 0 Then
                If IsNumeric(ws.Cells(r, amtCol).Value) Then amtVal = CDbl(ws.Cells(r, amtCol).Value)
            End If
            
            accNo = "": descText = "": cpName = ""
            If accCol > 0 Then accNo = CStr(ws.Cells(r, accCol).Value)
            If descCol > 0 Then descText = CStr(ws.Cells(r, descCol).Value)
            If cpCol > 0 Then cpName = CStr(ws.Cells(r, cpCol).Value)
            
            ReDim recArray(0 To 6)
            recArray(0) = txID
            recArray(1) = isAlert
            recArray(2) = IIf(IsEmpty(parsedDate), dtVal, parsedDate)
            recArray(3) = amtVal
            recArray(4) = accNo
            recArray(5) = descText
            recArray(6) = cpName
            
            If Not DictExists(dict, txID) Then
                DictAdd dict, txID, recArray
            End If
        End If
    Next r
    
    Set ExtractTransactions = dict
End Function

' =========================================================================
' [AUDIT] 11. GROUPED & PERFECTLY ALIGNED DISCREPANCY AUDIT WRITER
' =========================================================================
Private Sub RecordAddedTransactions(ByVal wsAdded As Worksheet, ByVal ecmID As String, ByVal alertID As String, _
                                    ByVal newMap As Object, ByVal oldMap As Object, _
                                    ByVal newCPList As Collection, ByRef addedCount As Long, ByRef alertedAdded As Long)
    Dim txKeys As Variant
    Dim i As Long, curRow As Long
    Dim txID As String, cpName As String
    Dim rec As Variant
    Dim addedList As New Collection
    Dim isNewCP As Boolean, cpItem As Variant
    
    addedCount = 0
    alertedAdded = 0
    
    txKeys = DictKeys(newMap)
    For i = LBound(txKeys) To UBound(txKeys)
        txID = CStr(txKeys(i))
        If Not DictExists(oldMap, txID) Then
            rec = DictGet(newMap, txID)
            addedCount = addedCount + 1
            If CStr(rec(1)) = "Yes" Then alertedAdded = alertedAdded + 1
            addedList.Add rec
        End If
    Next i
    
    ' If no transactions added for this alert, do not pollute the audit sheet
    If addedCount = 0 Then Exit Sub
    
    curRow = wsAdded.Cells(wsAdded.Rows.Count, 2).End(xlUp).Row + 1
    If curRow < 4 Then curRow = 4 Else curRow = curRow + 2 ' clean 2-row gap between alerts
    
    ' 1. Distinct Alert Header Banner (Spanning Columns B to I - Exactly 8 Columns)
    wsAdded.Range("B" & curRow & ":I" & curRow).Merge
    With wsAdded.Cells(curRow, 2)
        .Value = "ECM ID: " & ecmID & "   |   ALERT ID: " & alertID & "   |   NEW TRANSACTIONS ADDED: " & addedCount & _
                 "  (" & alertedAdded & " Alerted [ALERT], " & (addedCount - alertedAdded) & " Non-Alerted)" & _
                 IIf(newCPList.Count > 0, "   |   NEW PIVOT CPs: " & newCPList.Count, "")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    wsAdded.Rows(curRow).RowHeight = 26
    curRow = curRow + 1
    
    ' 2. Pivot Counterparty Callout Bar (if newly added counterparties exist in Pivot)
    If newCPList.Count > 0 Then
        wsAdded.Range("B" & curRow & ":I" & curRow).Merge
        With wsAdded.Cells(curRow, 2)
            .Value = "[!] PIVOT RECONCILIATION: " & newCPList.Count & " New Counterparty Added in Expanded Lookback -> " & CStr(newCPList(1))
            .Font.Name = "Segoe UI"
            .Font.Size = 9.5
            .Font.Bold = True
            .Font.Color = RGB(180, 83, 9)       ' Amber 800
            .Interior.Color = RGB(254, 243, 199) ' Amber 100
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        With wsAdded.Range("B" & curRow & ":I" & curRow).Borders
            .LineStyle = xlContinuous
            .Color = RGB(251, 191, 36)
            .Weight = xlThin
        End With
        wsAdded.Rows(curRow).RowHeight = 22
        curRow = curRow + 1
    End If
    
    ' 3. Column Sub-headers (Columns B to I - Exactly 8 Columns)
    Dim subHeaders As Variant
    subHeaders = Array("Transaction ID", "Is Alerted?", "Transaction Date", "Amount ($)", "Account No", "Transaction Description", "Counterparty Name", "Pivot Scope Status")
    Dim cIdx As Long
    For cIdx = 0 To UBound(subHeaders)
        With wsAdded.Cells(curRow, cIdx + 2)
            .Value = subHeaders(cIdx)
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = RGB(51, 65, 85) ' Slate 700
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next cIdx
    wsAdded.Rows(curRow).RowHeight = 22
    curRow = curRow + 1
    
    ' 4. List of Transactions for this Alert
    Dim itemVar As Variant
    For Each itemVar In addedList
        rec = itemVar
        cpName = Trim(CStr(rec(6)))
        
        ' Check if counterparty was newly added in Pivot
        isNewCP = False
        For Each cpItem In newCPList
            If UCase(Trim(CStr(cpItem))) = UCase(cpName) Or InStr(1, UCase(cpName), UCase(Trim(CStr(cpItem))), vbTextCompare) > 0 Then
                isNewCP = True
                Exit For
            End If
        Next cpItem
        
        wsAdded.Cells(curRow, 2).Value = rec(0) ' Txn ID
        wsAdded.Cells(curRow, 3).Value = IIf(CStr(rec(1)) = "Yes", "Yes [ALERT]", "No")
        wsAdded.Cells(curRow, 4).Value = rec(2) ' Date
        wsAdded.Cells(curRow, 5).Value = rec(3) ' Amount
        wsAdded.Cells(curRow, 6).Value = rec(4) ' Account No
        wsAdded.Cells(curRow, 7).Value = rec(5) ' Description
        wsAdded.Cells(curRow, 8).Value = rec(6) ' Counterparty
        wsAdded.Cells(curRow, 9).Value = IIf(isNewCP, "NEW COUNTERPARTY [ALERT]", "Existing Scope [OK]")
        
        wsAdded.Cells(curRow, 2).HorizontalAlignment = xlCenter
        wsAdded.Cells(curRow, 3).HorizontalAlignment = xlCenter
        wsAdded.Cells(curRow, 4).HorizontalAlignment = xlCenter
        If IsDate(rec(2)) Then wsAdded.Cells(curRow, 4).NumberFormat = "yyyy-mm-dd"
        wsAdded.Cells(curRow, 5).HorizontalAlignment = xlRight
        wsAdded.Cells(curRow, 5).NumberFormat = "0,##0.00"
        wsAdded.Cells(curRow, 6).HorizontalAlignment = xlLeft
        wsAdded.Cells(curRow, 7).HorizontalAlignment = xlLeft
        wsAdded.Cells(curRow, 8).HorizontalAlignment = xlLeft
        wsAdded.Cells(curRow, 9).HorizontalAlignment = xlCenter
        
        wsAdded.Rows(curRow).RowHeight = 20
        
        With wsAdded.Range("B" & curRow & ":I" & curRow).Borders
            .LineStyle = xlContinuous
            .Color = COLOR_BORDER
            .Weight = xlThin
        End With
        
        ' Conditional styling for Alerted transaction
        If CStr(rec(1)) = "Yes" Then
            wsAdded.Cells(curRow, 3).Font.Bold = True
            wsAdded.Cells(curRow, 3).Font.Color = COLOR_ALERT_RED
            wsAdded.Cells(curRow, 3).Interior.Color = RGB(254, 242, 242)
        Else
            wsAdded.Cells(curRow, 3).Font.Color = RGB(71, 85, 105)
        End If
        
        ' Conditional styling for Newly added Counterparty
        If isNewCP Then
            wsAdded.Cells(curRow, 8).Font.Bold = True
            wsAdded.Cells(curRow, 9).Font.Bold = True
            wsAdded.Cells(curRow, 9).Font.Color = COLOR_ALERT_RED
            wsAdded.Cells(curRow, 9).Interior.Color = RGB(254, 242, 242)
        Else
            wsAdded.Cells(curRow, 9).Font.Color = COLOR_SUCCESS_GREEN
        End If
        
        curRow = curRow + 1
    Next itemVar
End Sub

Private Sub RecordDroppedTransactions(ByVal wsDropped As Worksheet, ByVal ecmID As String, ByVal alertID As String, _
                                      ByVal oldMap As Object, ByVal newMap As Object, _
                                      ByRef droppedCount As Long)
    Dim txKeys As Variant
    Dim i As Long, curRow As Long
    Dim txID As String
    Dim rec As Variant
    Dim droppedList As New Collection
    
    droppedCount = 0
    txKeys = DictKeys(oldMap)
    For i = LBound(txKeys) To UBound(txKeys)
        txID = CStr(txKeys(i))
        If Not DictExists(newMap, txID) Then
            rec = DictGet(oldMap, txID)
            droppedCount = droppedCount + 1
            droppedList.Add rec
        End If
    Next i
    
    If droppedCount = 0 Then Exit Sub
    
    curRow = wsDropped.Cells(wsDropped.Rows.Count, 2).End(xlUp).Row + 1
    If curRow < 4 Then curRow = 4 Else curRow = curRow + 2
    
    wsDropped.Range("B" & curRow & ":H" & curRow).Merge
    With wsDropped.Cells(curRow, 2)
        .Value = "ECM ID: " & ecmID & "  |  ALERT ID: " & alertID & "  |  DROPPED TRANSACTIONS: " & droppedCount
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    wsDropped.Rows(curRow).RowHeight = 26
    curRow = curRow + 1
    
    Dim subHeaders As Variant
    subHeaders = Array("Transaction ID", "Is Alerted?", "Transaction Date", "Amount ($)", "Account No", "Transaction Description", "Counterparty Name")
    Dim cIdx As Long
    For cIdx = 0 To UBound(subHeaders)
        With wsDropped.Cells(curRow, cIdx + 2)
            .Value = subHeaders(cIdx)
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = RGB(51, 65, 85)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next cIdx
    wsDropped.Rows(curRow).RowHeight = 22
    curRow = curRow + 1
    
    Dim itemVar As Variant
    For Each itemVar In droppedList
        rec = itemVar
        wsDropped.Cells(curRow, 2).Value = rec(0)
        wsDropped.Cells(curRow, 3).Value = rec(1)
        wsDropped.Cells(curRow, 4).Value = rec(2)
        wsDropped.Cells(curRow, 5).Value = rec(3)
        wsDropped.Cells(curRow, 6).Value = rec(4)
        wsDropped.Cells(curRow, 7).Value = rec(5)
        wsDropped.Cells(curRow, 8).Value = rec(6)
        
        wsDropped.Cells(curRow, 2).HorizontalAlignment = xlCenter
        wsDropped.Cells(curRow, 3).HorizontalAlignment = xlCenter
        wsDropped.Cells(curRow, 4).HorizontalAlignment = xlCenter
        If IsDate(rec(2)) Then wsDropped.Cells(curRow, 4).NumberFormat = "yyyy-mm-dd"
        wsDropped.Cells(curRow, 5).HorizontalAlignment = xlRight
        wsDropped.Cells(curRow, 5).NumberFormat = "0,##0.00"
        wsDropped.Cells(curRow, 6).HorizontalAlignment = xlLeft
        wsDropped.Cells(curRow, 7).HorizontalAlignment = xlLeft
        wsDropped.Cells(curRow, 8).HorizontalAlignment = xlLeft
        
        wsDropped.Rows(curRow).RowHeight = 20
        With wsDropped.Range("B" & curRow & ":H" & curRow).Borders
            .LineStyle = xlContinuous
            .Color = COLOR_BORDER
            .Weight = xlThin
        End With
        curRow = curRow + 1
    Next itemVar
End Sub

' =========================================================================
' [UI] 12. DASHBOARD FORMATTING & BUTTON SETUP (100% ALIGNED)
' =========================================================================
Public Function Setup_Comparison_Dashboard(ByVal wb As Workbook, Optional ByVal forceReset As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Dim shp As Shape
    
    Set ws = GetOrCreateWorksheet(wb, SHEET_DASHBOARD, True)
    
    If Not forceReset And ws.Range("B4").Value = "AML TRANSACTION MONITORING - LOOKBACK PERIOD BATCH RECONCILIATION COCKPIT" Then
        Set Setup_Comparison_Dashboard = ws
        Exit Function
    End If
    
    ws.Cells.Clear
    
    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0
    
    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo 0
    
    ' Exact Classic Column Widths (Columns B to O - 14 Columns total)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 14 ' ECM ID
    ws.Columns("C").ColumnWidth = 16 ' Alert ID
    ws.Columns("D").ColumnWidth = 32 ' Old File Name
    ws.Columns("E").ColumnWidth = 32 ' New File Name
    ws.Columns("F").ColumnWidth = 22 ' Old Period
    ws.Columns("G").ColumnWidth = 22 ' New Period
    ws.Columns("H").ColumnWidth = 11 ' Old Count
    ws.Columns("I").ColumnWidth = 11 ' New Count
    ws.Columns("J").ColumnWidth = 11 ' Count Delta
    ws.Columns("K").ColumnWidth = 11 ' % Change
    ws.Columns("L").ColumnWidth = 11 ' Old Alerted
    ws.Columns("M").ColumnWidth = 11 ' New Alerted
    ws.Columns("N").ColumnWidth = 11 ' Alerted Delta
    ws.Columns("O").ColumnWidth = 36 ' Status Badge
    
    ws.Rows("1:3").RowHeight = 12
    ws.Rows(4).RowHeight = 32
    ws.Rows(5).RowHeight = 20
    ws.Rows(6).RowHeight = 10
    ws.Rows("7:10").RowHeight = 22 ' KPI Cards
    ws.Rows(11).RowHeight = 12
    ws.Rows(12).RowHeight = 30     ' Button Bar
    ws.Rows(13).RowHeight = 12
    ws.Rows(14).RowHeight = 26     ' Table Header
    
    ' Master Banner Header (Columns B to O)
    ws.Range("B4:O4").Merge
    With ws.Range("B4")
        .Value = "AML TRANSACTION MONITORING - LOOKBACK PERIOD BATCH RECONCILIATION COCKPIT"
        .Font.Name = "Segoe UI"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Range("B5:O5").Merge
    With ws.Range("B5")
        .Value = "Automated discrepancy verification when review period dates expand | Quantifies transaction counts, newly alerted transactions, and Pivot counterparties"
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Color = RGB(100, 116, 139)
        .Interior.Color = RGB(241, 245, 249)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' 5 Balanced KPI Tiles (Columns B to O - Classic Version 3.0 Spans)
    CreateKPITile ws, "B7:D9", "ALERTS EVALUATED", "0", "Total compared alert files"
    CreateKPITile ws, "E7:G9", "ALERTS WITH INCREASED COUNT", "0", "Lookback count increased [WARN]"
    CreateKPITile ws, "H7:J9", "TOTAL TXNS ADDED", "0", "Net transaction volume growth"
    CreateKPITile ws, "K7:M9", "NEW ALERTED TXNS", "0", "Newly captured in-scope [ALERT]"
    CreateKPITile ws, "N7:O9", "ALERTS UNCHANGED", "0", "Exact count match [OK]"
    
    ' 5 Action Buttons Perfectly Aligned with KPI Cards Directly Above (Row 12 across B to O)
    CreateActionButton ws, 12, "B", "D", "[RUN] Batch Comparison", "Run_Batch_Comparison", COLOR_ACCENT_BLUE
    CreateActionButton ws, 12, "E", "G", "[OPEN] Compare Single Pair", "Run_Single_Pair_Comparison", RGB(79, 70, 229)
    CreateActionButton ws, 12, "H", "J", "[EXPORT] Export to Workbook", "Export_To_New_Workbook", RGB(16, 149, 193)
    CreateActionButton ws, 12, "K", "M", "[VIEW] View Added Txns", "View_Added_Transactions", RGB(13, 148, 136)
    CreateActionButton ws, 12, "N", "O", "[RESET] Reset Dashboard", "Clear_Comparison_Dashboard", RGB(100, 116, 139)
    
    ' Table Header Row (Columns B to O - All 14 Columns from Version 3.0)
    Dim headers As Variant
    headers = Array("ECM ID", "Alert ID", "Old File Name", "New File Name", "Old Date Period", "New Date Period", _
                    "Old Count", "New Count", "Count Delta", "% Change", "Old Alerted", "New Alerted", "Alerted Delta", "Reconciliation Status")
                    
    Dim c As Long
    For c = 0 To UBound(headers)
        With ws.Cells(14, c + 2)
            .Value = headers(c)
            .Font.Name = "Segoe UI"
            .Font.Size = 9.5
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = COLOR_HEADER_BG
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next c
    
    Set Setup_Comparison_Dashboard = ws
End Function

Private Sub CreateKPITile(ByVal ws As Worksheet, ByVal cellRange As String, ByVal title As String, ByVal valStr As String, ByVal subTitle As String)
    With ws.Range(cellRange)
        .Interior.Color = COLOR_CARD_BG
        With .Borders
            .LineStyle = xlContinuous
            .Color = COLOR_BORDER
            .Weight = xlThin
        End With
    End With
    
    Dim firstCell As Range
    Set firstCell = ws.Range(cellRange).Cells(1, 1)
    
    With firstCell
        .Value = title
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Font.Bold = True
        .Font.Color = RGB(100, 116, 139)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With firstCell.Offset(1, 0)
        .Value = valStr
        .Font.Name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(15, 23, 42)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With firstCell.Offset(2, 0)
        .Value = subTitle
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(148, 163, 184)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub UpdateKPICards(ByVal ws As Worksheet, ByVal totalAlerts As Long, ByVal increasedCount As Long, _
                           ByVal addedTxns As Long, ByVal newAlerted As Long, ByVal unchangedCount As Long)
    ws.Range("B8").Value = totalAlerts
    ws.Range("E8").Value = increasedCount
    If increasedCount > 0 Then ws.Range("E8").Font.Color = COLOR_AMBER
    
    ws.Range("H8").Value = addedTxns
    If addedTxns > 0 Then ws.Range("H8").Font.Color = COLOR_ACCENT_BLUE
    
    ws.Range("K8").Value = newAlerted
    If newAlerted > 0 Then ws.Range("K8").Font.Color = COLOR_ALERT_RED
    
    ws.Range("N8").Value = unchangedCount
    If unchangedCount > 0 Then ws.Range("N8").Font.Color = COLOR_SUCCESS_GREEN
End Sub

Private Sub CreateActionButton(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal startCol As String, ByVal endCol As String, _
                               ByVal btnText As String, ByVal macroName As String, ByVal bgColor As Long)
    Dim btnRange As Range
    Dim shp As Shape
    
    Set btnRange = ws.Range(startCol & rowIdx & ":" & endCol & rowIdx)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnRange.Left + 2, btnRange.Top + 2, btnRange.Width - 4, btnRange.Height - 4)
    
    With shp
        .Name = "Btn_" & macroName
        .OnAction = macroName
        .Adjustments.Item(1) = 0.2
        .Fill.Solid
        .Fill.ForeColor.RGB = bgColor
        .Line.Visible = msoFalse
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = btnText
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Private Sub FormatBatchTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    If endRow < startRow Then Exit Sub
    
    Dim rngTable As Range
    Set rngTable = ws.Range("B" & startRow & ":O" & endRow)
    
    With rngTable
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous
            .Color = COLOR_BORDER
            .Weight = xlThin
        End With
    End With
    
    Dim r As Long
    For r = startRow To endRow
        ws.Rows(r).RowHeight = 22
        
        ws.Cells(r, 2).HorizontalAlignment = xlCenter ' ECM ID
        ws.Cells(r, 3).HorizontalAlignment = xlCenter ' Alert ID
        ws.Cells(r, 4).HorizontalAlignment = xlLeft   ' Old File
        ws.Cells(r, 5).HorizontalAlignment = xlLeft   ' New File
        ws.Cells(r, 6).HorizontalAlignment = xlCenter ' Old Period
        ws.Cells(r, 7).HorizontalAlignment = xlCenter ' New Period
        
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 10)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 10)).NumberFormat = "#,##0"
        
        ws.Cells(r, 11).HorizontalAlignment = xlRight
        ws.Cells(r, 11).NumberFormat = "0.00%"
        
        ws.Range(ws.Cells(r, 12), ws.Cells(r, 14)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(r, 12), ws.Cells(r, 14)).NumberFormat = "#,##0"
        
        If ws.Cells(r, 10).Value > 0 Then
            ws.Cells(r, 10).Font.Bold = True
            ws.Cells(r, 10).Font.Color = COLOR_ALERT_RED
            ws.Cells(r, 10).Interior.Color = RGB(254, 242, 242)
        End If
        
        If ws.Cells(r, 14).Value > 0 Then
            ws.Cells(r, 14).Font.Bold = True
            ws.Cells(r, 14).Font.Color = COLOR_ALERT_RED
        End If
        
        With ws.Cells(r, 15)
            .HorizontalAlignment = xlLeft
            .Font.Bold = True
            If InStr(.Value, "NEW ALERTED") > 0 Then
                .Font.Color = COLOR_ALERT_RED
                .Interior.Color = RGB(254, 242, 242)
            ElseIf InStr(.Value, "NEW CP") > 0 Then
                .Font.Color = COLOR_AMBER
                .Interior.Color = RGB(254, 243, 199)
            ElseIf InStr(.Value, "COUNT INCREASED") > 0 Then
                .Font.Color = COLOR_AMBER
                .Interior.Color = RGB(254, 243, 199)
            ElseIf InStr(.Value, "NO COUNT CHANGE") > 0 Then
                .Font.Color = COLOR_SUCCESS_GREEN
                .Interior.Color = RGB(240, 253, 244)
            Else
                .Font.Color = RGB(71, 85, 105)
            End If
        End With
    Next r
End Sub

Private Sub WriteUnmatchedRow(ByVal wsDash As Worksheet, ByVal rowIdx As Long, ByVal ecmID As String, ByVal alertID As String, _
                              ByVal oldF As String, ByVal newF As String, ByVal statusStr As String)
    wsDash.Cells(rowIdx, 2).Value = ecmID
    wsDash.Cells(rowIdx, 3).Value = alertID
    wsDash.Cells(rowIdx, 4).Value = oldF
    wsDash.Cells(rowIdx, 5).Value = newF
    wsDash.Cells(rowIdx, 6).Value = "N/A"
    wsDash.Cells(rowIdx, 7).Value = "N/A"
    wsDash.Cells(rowIdx, 8).Value = "-"
    wsDash.Cells(rowIdx, 9).Value = "-"
    wsDash.Cells(rowIdx, 10).Value = "-"
    wsDash.Cells(rowIdx, 11).Value = "-"
    wsDash.Cells(rowIdx, 12).Value = "-"
    wsDash.Cells(rowIdx, 13).Value = "-"
    wsDash.Cells(rowIdx, 14).Value = "-"
    wsDash.Cells(rowIdx, 15).Value = statusStr
    wsDash.Cells(rowIdx, 15).Font.Bold = True
    wsDash.Cells(rowIdx, 15).Font.Color = COLOR_ALERT_RED
End Sub

Private Sub FormatAuditSheetHeaders(ByVal ws As Worksheet, ByVal bannerTitle As String, Optional ByVal isAdded As Boolean = True)
    ws.Cells.Clear
    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo 0
    
    ApplyAuditSheetColumnWidths ws, isAdded
    
    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 10
    
    ' Master Banner Header
    If isAdded Then
        ws.Range("B1:I1").Merge
        With ws.Range("B1")
            .Value = bannerTitle
            .Font.Name = "Segoe UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = COLOR_HEADER_BG
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Else
        ws.Range("B1:H1").Merge
        With ws.Range("B1")
            .Value = bannerTitle
            .Font.Name = "Segoe UI"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = COLOR_HEADER_BG
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
End Sub

Private Sub ApplyAuditSheetColumnWidths(ByVal ws As Worksheet, Optional ByVal isAdded As Boolean = True)
    If ws Is Nothing Then Exit Sub
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 18 ' Txn ID
    ws.Columns("C").ColumnWidth = 14 ' Alerted?
    ws.Columns("D").ColumnWidth = 14 ' Date
    ws.Columns("E").ColumnWidth = 16 ' Amount ($)
    ws.Columns("F").ColumnWidth = 24 ' Account No
    ws.Columns("G").ColumnWidth = 28 ' Description
    ws.Columns("H").ColumnWidth = 32 ' Counterparty
    If isAdded Then
        ws.Columns("I").ColumnWidth = 26 ' Pivot Scope Status
    End If
End Sub

' =========================================================================
' [HELPER] 13. CROSS-PLATFORM FILE & STRING HELPERS
' =========================================================================
Private Function Pick_Folder(ByVal promptTitle As String) As String
#If Mac Then
    Dim script As String, result As String
    script = "return POSIX path of (choose folder with prompt """ & promptTitle & """)"
    On Error Resume Next
    result = MacScript(script)
    On Error GoTo 0
    Pick_Folder = Trim(result)
#Else
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    fd.Title = promptTitle
    If fd.Show = -1 Then
        Pick_Folder = fd.SelectedItems(1)
    Else
        Pick_Folder = ""
    End If
#End If
End Function

Private Function Pick_Excel_File(ByVal promptTitle As String) As String
    Dim vFile As Variant
    vFile = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , promptTitle)
    If VarType(vFile) = vbBoolean And vFile = False Then
        Pick_Excel_File = ""
    Else
        Pick_Excel_File = CStr(vFile)
    End If
End Function

Private Function ListExcelFiles(ByVal folderPath As String) As Collection
    Dim col As New Collection
    Dim sFile As String, pSep As String
    
    pSep = Application.PathSeparator
    If Right(folderPath, 1) <> pSep Then folderPath = folderPath & pSep
    
    sFile = Dir(folderPath & "*.xlsx")
    Do While sFile <> ""
        If Left(sFile, 2) <> "~$" Then col.Add folderPath & sFile
        sFile = Dir()
    Loop
    
    sFile = Dir(folderPath & "*.xlsm")
    Do While sFile <> ""
        If Left(sFile, 2) <> "~$" Then col.Add folderPath & sFile
        sFile = Dir()
    Loop
    
    Set ListExcelFiles = col
End Function

' Extract ECM ID (e.g. 138896 from 138896_ALERT2365101)
Private Function ExtractECMID(ByVal fPath As String) As String
    Dim fName As String
    Dim pos As Long, pre As String, i As Long, ch As String, res As String
    fName = GetFileName(fPath)
    
    pos = InStr(1, UCase(fName), "ALERT")
    If pos > 1 Then
        pre = Trim(Left(fName, pos - 1))
        Do While Right(pre, 1) = "_" Or Right(pre, 1) = "-" Or Right(pre, 1) = " "
            pre = Left(pre, Len(pre) - 1)
        Loop
        For i = Len(pre) To 1 Step -1
            ch = Mid(pre, i, 1)
            If ch >= "0" And ch <= "9" Then
                res = ch & res
            Else
                If Len(res) > 0 Then Exit For
            End If
        Next i
        If Len(res) > 0 Then
            ExtractECMID = res
            Exit Function
        End If
    End If
    ExtractECMID = "-"
End Function

' Extract Alert ID (e.g. ALERT2365101)
Private Function ExtractAlertID(ByVal fPath As String) As String
    Dim fName As String
    Dim pos As Long, i As Long, ch As String, res As String
    fName = GetFileName(fPath)
    
    pos = InStr(1, UCase(fName), "ALERT")
    If pos > 0 Then
        res = "ALERT"
        For i = pos + 5 To Len(fName)
            ch = Mid(fName, i, 1)
            If ch >= "0" And ch <= "9" Then
                res = res & ch
            Else
                Exit For
            End If
        Next i
        ExtractAlertID = res
        Exit Function
    End If
    
    pos = InStrRev(fName, ".")
    If pos > 0 Then ExtractAlertID = Left(fName, pos - 1) Else ExtractAlertID = fName
End Function

Private Function ExtractMatchKey(ByVal fPath As String) As String
    Dim aID As String, eID As String
    aID = ExtractAlertID(fPath)
    eID = ExtractECMID(fPath)
    If aID <> "" And InStr(aID, "ALERT") > 0 Then
        ExtractMatchKey = aID
    ElseIf eID <> "" And eID <> "-" Then
        ExtractMatchKey = eID
    Else
        ExtractMatchKey = GetFileName(fPath)
    End If
End Function

Private Function ExtractDateFromFilename(ByVal fName As String) As String
    Dim posTo As Long, leftPart As String, rightPart As String
    Dim d1 As String, d2 As String
    
    posTo = InStr(1, LCase(fName), " to ")
    If posTo > 10 Then
        leftPart = Trim(Left(fName, posTo - 1))
        d1 = Right(leftPart, 10)
        If (Mid(d1, 3, 1) = "." Or Mid(d1, 3, 1) = "/" Or Mid(d1, 3, 1) = "-") And _
           (Mid(d1, 6, 1) = "." Or Mid(d1, 6, 1) = "/" Or Mid(d1, 6, 1) = "-") Then
            d1 = Replace(Replace(d1, ".", "-"), "/", "-")
            rightPart = Trim(Mid(fName, posTo + 4))
            If Len(rightPart) >= 10 Then
                d2 = Left(rightPart, 10)
                If (Mid(d2, 3, 1) = "." Or Mid(d2, 3, 1) = "/" Or Mid(d2, 3, 1) = "-") And _
                   (Mid(d2, 6, 1) = "." Or Mid(d2, 6, 1) = "/" Or Mid(d2, 6, 1) = "-") Then
                    d2 = Replace(Replace(d2, ".", "-"), "/", "-")
                    ExtractDateFromFilename = d1 & " to " & d2
                    Exit Function
                End If
            End If
        End If
    End If
    ExtractDateFromFilename = ""
End Function

Private Function ParseDateValue(ByVal val As Variant) As Variant
    If IsEmpty(val) Then Exit Function
    Dim s As String, d As Double
    s = Trim(CStr(val))
    If s = "" Then Exit Function
    
    If IsNumeric(s) Then
        d = Val(s)
        If d > 20000 And d < 70000 Then
            On Error Resume Next
            ParseDateValue = CDate(d)
            On Error GoTo 0
            Exit Function
        End If
    End If
    
    If InStr(s, ".") > 0 Then s = Replace(s, ".", "/")
    
    If IsDate(s) Then
        On Error Resume Next
        ParseDateValue = CDate(s)
        On Error GoTo 0
        Exit Function
    End If
End Function

Private Function GetFileName(ByVal fPath As String) As String
    Dim pSep As String, pos As Long
    pSep = Application.PathSeparator
    pos = InStrRev(fPath, pSep)
    If pos > 0 Then GetFileName = Mid(fPath, pos + 1) Else GetFileName = fPath
End Function

Private Function FindTransactionSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, sName As String
    
    For Each ws In wb.Worksheets
        sName = LCase(ws.Name)
        If InStr(sName, "lookback") > 0 Or InStr(sName, "raw transactions") > 0 Or InStr(sName, "inscope") > 0 Then
            Set FindTransactionSheet = ws
            Exit Function
        End If
    Next ws
    
    For Each ws In wb.Worksheets
        If ws.Cells(1, 1).Value <> "" Then
            If InStr(LCase(CStr(ws.Cells(1, 1).Value)), "transaction") > 0 Then
                Set FindTransactionSheet = ws
                Exit Function
            End If
        End If
    Next ws
    
    For Each ws In wb.Worksheets
        sName = LCase(ws.Name)
        If InStr(sName, "pivot") = 0 And InStr(sName, "dashboard") = 0 Then
            Set FindTransactionSheet = ws
            Exit Function
        End If
    Next ws
    
    Set FindTransactionSheet = wb.Worksheets(1)
End Function

Private Function FormatDateRange(ByVal dMin As Variant, ByVal dMax As Variant) As String
    If IsDate(dMin) And IsDate(dMax) Then
        FormatDateRange = Format(dMin, "yyyy-mm-dd") & " to " & Format(dMax, "yyyy-mm-dd")
    Else
        FormatDateRange = "N/A"
    End If
End Function

Private Function GetOrCreateWorksheet(ByVal wb As Workbook, ByVal sheetName As String, Optional ByVal isFirst As Boolean = False) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        If isFirst Then
            Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        End If
        ws.Name = sheetName
    End If
    
    ws.Visible = xlSheetVisible
    Set GetOrCreateWorksheet = ws
End Function

' =========================================================================
' [DICT] 14. UNIVERSAL KEY-VALUE LOOKUP
' =========================================================================
Private Function CreateLookupDict() As Object
    On Error Resume Next
    Set CreateLookupDict = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    
    If CreateLookupDict Is Nothing Then
        Set CreateLookupDict = New Collection
    End If
End Function

Private Function DictExists(ByVal d As Object, ByVal k As String) As Boolean
    If TypeName(d) = "Dictionary" Then
        DictExists = d.Exists(k)
    Else
        On Error Resume Next
        Dim dummy As Variant
        dummy = d(k)
        DictExists = (Err.Number = 0)
        On Error GoTo 0
    End If
End Function

Private Sub DictAdd(ByVal d As Object, ByVal k As String, ByVal v As Variant)
    If TypeName(d) = "Dictionary" Then
        d.Add k, v
    Else
        d.Add v, k
    End If
End Sub

Private Function DictGet(ByVal d As Object, ByVal k As String) As Variant
    If TypeName(d) = "Dictionary" Then
        DictGet = d(k)
    Else
        DictGet = d(k)
    End If
End Function

Private Function DictCount(ByVal d As Object) As Long
    If d Is Nothing Then
        DictCount = 0
    ElseIf TypeName(d) = "Dictionary" Then
        DictCount = d.Count
    Else
        DictCount = d.Count
    End If
End Function

Private Function DictKeys(ByVal d As Object) As Variant
    If TypeName(d) = "Dictionary" Then
        DictKeys = d.Keys
    Else
        Dim keysArr() As String
        Dim count As Long, i As Long
        count = d.Count
        If count = 0 Then
            DictKeys = Array()
            Exit Function
        End If
        ReDim keysArr(0 To count - 1)
        For i = 1 To count
            Dim item As Variant
            item = d(i)
            If IsArray(item) Then
                keysArr(i - 1) = CStr(item(0))
            Else
                keysArr(i - 1) = CStr(item)
            End If
        Next i
        DictKeys = keysArr
    End If
End Function
