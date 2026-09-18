Attribute VB_Name = "AML_Period_Comparison"
' =========================================================================
' [AML] TRANSACTION MONITORING: LOOKBACK PERIOD BATCH RECONCILIATION COCKPIT
' =========================================================================
' Version: 4.0 (Pivot Diff, Exact Schema Formatting, 1-Click Export & Cleaned Layout)
'
' Key Capabilities:
'  1. PIVOT COUNTERPARTY COMPARISON: Automatically compares Pivot tables between
'     Old and New files to identify newly added counterparties (e.g. MONICA HERZBERG
'     CLAUDIO EULAU) and links them to associated added transactions.
'  2. EXACT NEW FILE FORMATTING PRESERVED: Extra transactions identified retain the
'     exact column schema (all 44 columns) and original formatting (fonts, colors,
'     number formats, dates, amounts) directly from the new transaction file.
'  3. STREAMLINED DASHBOARD TABLE: Removed the % Change (delta change) column so
'     compliance review is 100% focused on absolute counts and alerted shifts.
'  4. ONE-CLICK EXPORT TO NEW WORKBOOK: Dedicated button exports a pristine,
'     formatted .xlsx workbook containing the Dashboard and all Discrepancy details.
'  5. SIMPLIFIED & GROUPED PER ALERT: Every alert is clearly differentiated with
'     bold header banners, individual Pivot reconciliation cards, and clean gaps.
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
    
    ' Copy Discrepancy Added Txns (with exact 44-col formatting and Pivot diff)
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
    Dim totalNewPivotCPs As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wsDash = wbDash.Worksheets(SHEET_DASHBOARD)
    Set wsAdded = GetOrCreateWorksheet(wbDash, SHEET_ADDED_TXNS)
    Set wsDropped = GetOrCreateWorksheet(wbDash, SHEET_DROPPED_TXNS)
    
    FormatAuditSheetHeaders wsAdded, "NEWLY ADDED TRANSACTIONS & PIVOT RECONCILIATION (EXPANDED LOOKBACK PERIOD)"
    FormatAuditSheetHeaders wsDropped, "DROPPED / EXCLUDED TRANSACTIONS GROUPED BY ALERT"
    
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
                        countIncreasedAlerts, countUnchangedAlerts, totalAddedTxns, totalDroppedTxns, _
                        totalAlertedAdded, totalNewPivotCPs
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
    UpdateKPICards wsDash, totalAlerts, countIncreasedAlerts, totalAddedTxns, totalAlertedAdded, totalNewPivotCPs
    
    wsAdded.Columns.AutoFit
    wsDropped.Columns.AutoFit
    wsDash.Activate
    
    MsgBox "Batch Comparison Complete!" & vbCrLf & vbCrLf & _
           "- Total Alerts Evaluated: " & totalAlerts & vbCrLf & _
           "- Alerts With Increased Count: " & countIncreasedAlerts & vbCrLf & _
           "- Total Extra Transactions Added: " & totalAddedTxns & vbCrLf & _
           "- New Alerted Transactions: " & totalAlertedAdded & vbCrLf & _
           "- New Counterparties Identified in Pivot: " & totalNewPivotCPs, vbInformation, "Reconciliation Finished"

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
    Dim totalNewPivotCPs As Long
    Dim curRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wsDash = wbDash.Worksheets(SHEET_DASHBOARD)
    Set wsAdded = GetOrCreateWorksheet(wbDash, SHEET_ADDED_TXNS)
    Set wsDropped = GetOrCreateWorksheet(wbDash, SHEET_DROPPED_TXNS)
    
    FormatAuditSheetHeaders wsAdded, "NEWLY ADDED TRANSACTIONS & PIVOT RECONCILIATION (EXPANDED LOOKBACK PERIOD)"
    FormatAuditSheetHeaders wsDropped, "DROPPED / EXCLUDED TRANSACTIONS GROUPED BY ALERT"
    
    ecmID = ExtractECMID(newFilePath)
    alertID = ExtractAlertID(newFilePath)
    curRow = 15
    
    wsDash.Range("B15:O30").ClearContents
    wsDash.Range("B15:O30").ClearFormats
    
    ProcessPair wsDash, wsAdded, wsDropped, curRow, ecmID, alertID, oldFilePath, newFilePath, _
                countIncreasedAlerts, countUnchangedAlerts, totalAddedTxns, totalDroppedTxns, _
                totalAlertedAdded, totalNewPivotCPs
                
    FormatBatchTable wsDash, 15, 15
    UpdateKPICards wsDash, 1, countIncreasedAlerts, totalAddedTxns, totalAlertedAdded, totalNewPivotCPs
    
    wsAdded.Columns.AutoFit
    wsDropped.Columns.AutoFit
    wsDash.Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Single Alert Comparison Complete!" & vbCrLf & vbCrLf & _
           "ECM ID: " & ecmID & " | Alert ID: " & alertID & vbCrLf & _
           "- Extra Transactions Identified: " & totalAddedTxns & vbCrLf & _
           "- New Alerted Transactions: " & totalAlertedAdded & vbCrLf & _
           "- New Counterparties Identified in Pivot: " & totalNewPivotCPs, vbInformation, "Reconciliation Finished"

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
' [PROCESS] 8. PROCESS INDIVIDUAL ALERT PAIR (WITH PIVOT DIFF & EXACT FORMATTING)
' =========================================================================
Private Sub ProcessPair(ByVal wsDash As Worksheet, ByVal wsAdded As Worksheet, ByVal wsDropped As Worksheet, _
                        ByVal rowIdx As Long, ByVal ecmID As String, ByVal alertID As String, _
                        ByVal oldFilePath As String, ByVal newFilePath As String, _
                        ByRef countIncreasedAlerts As Long, ByRef countUnchangedAlerts As Long, _
                        ByRef totalAddedTxns As Long, ByRef totalDroppedTxns As Long, _
                        ByRef totalAlertedAdded As Long, ByRef totalNewPivotCPs As Long)
                        
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
    
    ' 1. Open and extract Old File
    Set wbOld = Workbooks.Open(oldFilePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsOld = FindTransactionSheet(wbOld)
    Set oldMap = ExtractTransactions(wsOld, oldTxCount, oldAlertCount, oldMinDate, oldMaxDate)
    Set oldPivotCPs = ExtractPivotCounterparties(wbOld)
    
    ' 2. Open and extract New File
    Set wbNew = Workbooks.Open(newFilePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsNew = FindTransactionSheet(wbNew)
    Set newMap = ExtractTransactions(wsNew, newTxCount, newAlertCount, newMinDate, newMaxDate)
    Set newPivotCPs = ExtractPivotCounterparties(wbNew)
    
    ' 3. Compare Pivot Tables for New Counterparties Added
    Dim cpKeyVar As Variant, cpUpper As String
    For Each cpKeyVar In DictKeys(newPivotCPs)
        cpUpper = CStr(cpKeyVar)
        If Not DictExists(oldPivotCPs, cpUpper) Then
            newCPList.Add DictGet(newPivotCPs, cpUpper)
        End If
    Next cpKeyVar
    
    totalNewPivotCPs = totalNewPivotCPs + newCPList.Count
    
    ' 4. Date Resolution (Sheet cells with filename regex fallback)
    oldPeriodStr = FormatDateRange(oldMinDate, oldMaxDate)
    If oldPeriodStr = "N/A" Then oldPeriodStr = ExtractDateFromFilename(GetFileName(oldFilePath))
    If oldPeriodStr = "" Then oldPeriodStr = "N/A"
    
    newPeriodStr = FormatDateRange(newMinDate, newMaxDate)
    If newPeriodStr = "N/A" Then newPeriodStr = ExtractDateFromFilename(GetFileName(newFilePath))
    If newPeriodStr = "" Then newPeriodStr = "N/A"
    
    countDelta = newTxCount - oldTxCount
    alertDelta = newAlertCount - oldAlertCount
    
    If countDelta > 0 Then
        countIncreasedAlerts = countIncreasedAlerts + 1
        If alertDelta > 0 Then
            statusStr = "COUNT INCREASED (+" & countDelta & ") | NEW ALERTED (+" & alertDelta & ") [ALERT]"
        ElseIf newCPList.Count > 0 Then
            statusStr = "COUNT INCREASED (+" & countDelta & ") | NEW CP ADDED (" & newCPList.Count & ") [ALERT]"
        Else
            statusStr = "COUNT INCREASED (+" & countDelta & ") [WARN]"
        End If
    ElseIf countDelta < 0 Then
        statusStr = "COUNT DECREASED (" & countDelta & ") [-]"
    Else
        countUnchangedAlerts = countUnchangedAlerts + 1
        If newCPList.Count > 0 Then
            statusStr = "NO COUNT CHANGE (0) | NEW CP ADDED (" & newCPList.Count & ") [ALERT]"
        Else
            statusStr = "NO COUNT CHANGE (0) [OK]"
        End If
    End If
    
    ' 5. Write Dashboard Table Row (Columns B to O - 14 Columns, No % Change)
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
        .Cells(rowIdx, 11).Value = oldAlertCount
        .Cells(rowIdx, 12).Value = newAlertCount
        .Cells(rowIdx, 13).Value = alertDelta
        .Cells(rowIdx, 14).Value = IIf(newCPList.Count > 0, newCPList.Count & " New CP", "0")
        .Cells(rowIdx, 15).Value = statusStr
    End With
    
    ' 6. Record Added Transactions & Pivot Counterparty Diff into Discrepancy Sheet
    ' (wsNew remains open during copy to preserve 100% exact formatting and all 44 columns)
    Dim pairAddedCount As Long, pairAlertedAdded As Long
    RecordAddedTransactions wsAdded, ecmID, alertID, wsNew, newMap, oldMap, newCPList, pairAddedCount, pairAlertedAdded
    
    Dim pairDroppedCount As Long
    RecordDroppedTransactions wsDropped, ecmID, alertID, wsOld, oldMap, newMap, pairDroppedCount
    
    ' 7. Close workbooks safely
    Application.CutCopyMode = False
    wbOld.Close SaveChanges:=False
    wbNew.Close SaveChanges:=False
    
    totalAddedTxns = totalAddedTxns + pairAddedCount
    totalDroppedTxns = totalDroppedTxns + pairDroppedCount
    totalAlertedAdded = totalAlertedAdded + pairAlertedAdded
End Sub

' =========================================================================
' [PIVOT] 9. EXTRACT COUNTERPARTIES FROM PIVOT TABLE & SHEET
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
    
    ' Strategy 1: Excel Native PivotTable Object Model
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
    
    ' Strategy 2: Resilient Column A Cell Scanning (Starting row 4)
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
' [METRICS] 10. EXTRACT TRANSACTIONS FROM RAW/LOOKBACK SHEET
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
    
    ' Scan rows 1 to 3 to locate header columns
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
            
            ReDim recArray(0 To 7)
            recArray(0) = txID
            recArray(1) = isAlert
            recArray(2) = IIf(IsEmpty(parsedDate), dtVal, parsedDate)
            recArray(3) = amtVal
            recArray(4) = accNo
            recArray(5) = descText
            recArray(6) = cpName
            recArray(7) = r ' Source row index in sheet!
            
            If Not DictExists(dict, txID) Then
                DictAdd dict, txID, recArray
            End If
        End If
    Next r
    
    Set ExtractTransactions = dict
End Function

' =========================================================================
' [AUDIT] 11. RECORD ADDED TRANSACTIONS (EXACT 44-COL FORMATTING & PIVOT DIFF)
' =========================================================================
Private Sub RecordAddedTransactions(ByVal wsAdded As Worksheet, ByVal ecmID As String, ByVal alertID As String, _
                                    ByVal wsNew As Worksheet, ByVal newMap As Object, ByVal oldMap As Object, _
                                    ByVal newCPList As Collection, ByRef addedCount As Long, ByRef alertedAdded As Long)
    Dim txKeys As Variant
    Dim i As Long, curRow As Long, srcRow As Long
    Dim txID As String, endColLetter As String
    Dim rec As Variant
    Dim addedList As New Collection
    Dim lastCol As Long
    
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
    
    ' If no transactions added and no new counterparties in Pivot, do not clutter sheet
    If addedCount = 0 And newCPList.Count = 0 Then Exit Sub
    
    curRow = wsAdded.Cells(wsAdded.Rows.Count, 2).End(xlUp).Row + 1
    If curRow < 4 Then curRow = 4 Else curRow = curRow + 2 ' clean 2-row separation between alerts
    
    lastCol = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column
    If lastCol < 7 Then lastCol = 44
    endColLetter = Split(wsAdded.Cells(1, lastCol + 3).Address, "$")(1)
    
    ' 1. Distinct Alert Master Header Banner (Spanning Columns B to endColLetter)
    wsAdded.Range("B" & curRow & ":" & endColLetter & curRow).Merge
    With wsAdded.Cells(curRow, 2)
        .Value = "ECM ID: " & ecmID & "   |   ALERT ID: " & alertID & "   |   EXTRA TRANSACTIONS IDENTIFIED: " & addedCount & _
                 " (" & alertedAdded & " Alerted [ALERT], " & (addedCount - alertedAdded) & " Non-Alerted)   |   NEW PIVOT COUNTERPARTIES: " & newCPList.Count
        .Font.Name = "Segoe UI"
        .Font.Size = 10.5
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    wsAdded.Rows(curRow).RowHeight = 28
    curRow = curRow + 1
    
    ' 2. Pivot Counterparty Reconciliation Sub-Block
    If newCPList.Count > 0 Then
        wsAdded.Range("B" & curRow & ":F" & curRow).Merge
        With wsAdded.Cells(curRow, 2)
            .Value = "[!] PIVOT RECONCILIATION: " & newCPList.Count & " NEW COUNTERPARTY(IES) IDENTIFIED IN EXPANDED LOOKBACK PIVOT"
            .Font.Name = "Segoe UI"
            .Font.Size = 9.5
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = COLOR_AMBER
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        wsAdded.Rows(curRow).RowHeight = 22
        curRow = curRow + 1
        
        wsAdded.Cells(curRow, 2).Value = "Item #"
        wsAdded.Cells(curRow, 3).Value = "New Counterparty Name (From Pivot)"
        wsAdded.Cells(curRow, 4).Value = "Status in Old Scope"
        wsAdded.Cells(curRow, 5).Value = "Associated In-Scope Added Txn ID"
        wsAdded.Cells(curRow, 6).Value = "Counterparty Flag"
        With wsAdded.Range("B" & curRow & ":F" & curRow)
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = RGB(71, 85, 105)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        wsAdded.Rows(curRow).RowHeight = 20
        curRow = curRow + 1
        
        Dim cpIdx As Long, cpVal As Variant, assocTxn As String
        cpIdx = 1
        For Each cpVal In newCPList
            assocTxn = FindTxnForCounterparty(newMap, CStr(cpVal))
            wsAdded.Cells(curRow, 2).Value = cpIdx
            wsAdded.Cells(curRow, 3).Value = CStr(cpVal)
            wsAdded.Cells(curRow, 4).Value = "NOT IN OLD SCOPE"
            wsAdded.Cells(curRow, 5).Value = IIf(assocTxn <> "", assocTxn, "Lookback Scope")
            wsAdded.Cells(curRow, 6).Value = "NEW COUNTERPARTY [ALERT]"
            
            wsAdded.Cells(curRow, 2).HorizontalAlignment = xlCenter
            wsAdded.Cells(curRow, 3).HorizontalAlignment = xlLeft
            wsAdded.Cells(curRow, 4).HorizontalAlignment = xlCenter
            wsAdded.Cells(curRow, 5).HorizontalAlignment = xlCenter
            wsAdded.Cells(curRow, 6).HorizontalAlignment = xlCenter
            wsAdded.Cells(curRow, 6).Font.Bold = True
            wsAdded.Cells(curRow, 6).Font.Color = COLOR_ALERT_RED
            wsAdded.Cells(curRow, 6).Interior.Color = RGB(254, 242, 242)
            
            With wsAdded.Range("B" & curRow & ":F" & curRow).Borders
                .LineStyle = xlContinuous
                .Color = COLOR_BORDER
                .Weight = xlThin
            End With
            wsAdded.Rows(curRow).RowHeight = 20
            curRow = curRow + 1
            cpIdx = cpIdx + 1
        Next cpVal
    Else
        wsAdded.Range("B" & curRow & ":F" & curRow).Merge
        With wsAdded.Cells(curRow, 2)
            .Value = "[OK] PIVOT RECONCILIATION: No new counterparties identified in Pivot (All counterparties in expanded lookback existed in previous scope)"
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = COLOR_SUCCESS_GREEN
            .Interior.Color = RGB(240, 253, 244)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        With wsAdded.Range("B" & curRow & ":F" & curRow).Borders
            .LineStyle = xlContinuous
            .Color = RGB(187, 247, 208)
            .Weight = xlThin
        End With
        wsAdded.Rows(curRow).RowHeight = 22
        curRow = curRow + 1
    End If
    
    curRow = curRow + 1 ' 1-row gap before transactions table
    
    ' 3. Extra Transactions Table Header (Preserving Exact Columns & Formatting from wsNew)
    wsAdded.Range("B" & curRow & ":" & endColLetter & curRow).Merge
    With wsAdded.Cells(curRow, 2)
        .Value = "EXTRA TRANSACTIONS IDENTIFIED IN LOOKBACK PERIOD (PRESERVING EXACT NEW FILE SCHEMA & FORMATTING)"
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = RGB(51, 65, 85)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    wsAdded.Rows(curRow).RowHeight = 22
    curRow = curRow + 1
    
    ' Set up ECM ID & Alert ID in Columns B and C, then copy all source headers from wsNew
    wsAdded.Cells(curRow, 2).Value = "ECM ID"
    wsAdded.Cells(curRow, 3).Value = "Alert ID"
    With wsAdded.Range("B" & curRow & ":C" & curRow)
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = RGB(51, 65, 85)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Copy entire original header row from wsNew to wsAdded starting at Column D
    wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(1, lastCol)).Copy wsAdded.Cells(curRow, 4)
    wsAdded.Rows(curRow).RowHeight = wsNew.Rows(1).RowHeight
    If wsAdded.Rows(curRow).RowHeight < 22 Then wsAdded.Rows(curRow).RowHeight = 22
    curRow = curRow + 1
    
    ' 4. Copy each added transaction row directly from wsNew with full formatting
    Dim itemVar As Variant
    For Each itemVar In addedList
        rec = itemVar
        srcRow = CLng(rec(7))
        
        wsAdded.Cells(curRow, 2).Value = ecmID
        wsAdded.Cells(curRow, 3).Value = alertID
        wsAdded.Cells(curRow, 2).HorizontalAlignment = xlCenter
        wsAdded.Cells(curRow, 3).HorizontalAlignment = xlCenter
        
        ' Copy entire row range with 100% fidelity (values, formatting, formulas, borders)
        wsNew.Range(wsNew.Cells(srcRow, 1), wsNew.Cells(srcRow, lastCol)).Copy wsAdded.Cells(curRow, 4)
        wsAdded.Rows(curRow).RowHeight = 20
        
        If CStr(rec(1)) = "Yes" Then
            wsAdded.Cells(curRow, 2).Interior.Color = RGB(254, 242, 242)
            wsAdded.Cells(curRow, 3).Interior.Color = RGB(254, 242, 242)
            wsAdded.Cells(curRow, 2).Font.Bold = True
            wsAdded.Cells(curRow, 3).Font.Bold = True
        End If
        curRow = curRow + 1
    Next itemVar
End Sub

' =========================================================================
' [AUDIT] 12. RECORD DROPPED TRANSACTIONS
' =========================================================================
Private Sub RecordDroppedTransactions(ByVal wsDropped As Worksheet, ByVal ecmID As String, ByVal alertID As String, _
                                      ByVal wsOld As Worksheet, ByVal oldMap As Object, ByVal newMap As Object, _
                                      ByRef droppedCount As Long)
    Dim txKeys As Variant
    Dim i As Long, curRow As Long, srcRow As Long
    Dim txID As String, endColLetter As String
    Dim rec As Variant
    Dim droppedList As New Collection
    Dim lastCol As Long
    
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
    
    lastCol = wsOld.Cells(1, wsOld.Columns.Count).End(xlToLeft).Column
    If lastCol < 7 Then lastCol = 34
    endColLetter = Split(wsDropped.Cells(1, lastCol + 3).Address, "$")(1)
    
    wsDropped.Range("B" & curRow & ":" & endColLetter & curRow).Merge
    With wsDropped.Cells(curRow, 2)
        .Value = "ECM ID: " & ecmID & "  |  ALERT ID: " & alertID & "  |  DROPPED / EXCLUDED TRANSACTIONS: " & droppedCount
        .Font.Name = "Segoe UI"
        .Font.Size = 10.5
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    wsDropped.Rows(curRow).RowHeight = 28
    curRow = curRow + 1
    
    wsDropped.Cells(curRow, 2).Value = "ECM ID"
    wsDropped.Cells(curRow, 3).Value = "Alert ID"
    With wsDropped.Range("B" & curRow & ":C" & curRow)
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = RGB(51, 65, 85)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    wsOld.Range(wsOld.Cells(1, 1), wsOld.Cells(1, lastCol)).Copy wsDropped.Cells(curRow, 4)
    wsDropped.Rows(curRow).RowHeight = 22
    curRow = curRow + 1
    
    Dim itemVar As Variant
    For Each itemVar In droppedList
        rec = itemVar
        srcRow = CLng(rec(7))
        wsDropped.Cells(curRow, 2).Value = ecmID
        wsDropped.Cells(curRow, 3).Value = alertID
        wsDropped.Cells(curRow, 2).HorizontalAlignment = xlCenter
        wsDropped.Cells(curRow, 3).HorizontalAlignment = xlCenter
        
        wsOld.Range(wsOld.Cells(srcRow, 1), wsOld.Cells(srcRow, lastCol)).Copy wsDropped.Cells(curRow, 4)
        wsDropped.Rows(curRow).RowHeight = 20
        curRow = curRow + 1
    Next itemVar
End Sub

' Helper to match counterparty name to added transaction ID
Private Function FindTxnForCounterparty(ByVal newMap As Object, ByVal cpName As String) As String
    Dim k As Variant, rec As Variant
    Dim target As String
    target = UCase(Trim(cpName))
    
    For Each k In DictKeys(newMap)
        rec = DictGet(newMap, CStr(k))
        If UCase(Trim(CStr(rec(6)))) = target Then
            FindTxnForCounterparty = CStr(rec(0))
            Exit Function
        End If
    Next k
    FindTxnForCounterparty = ""
End Function

' =========================================================================
' [UI] 13. DASHBOARD FORMATTING & BUTTON SETUP
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
    
    ' Column Widths (Columns B to O - 14 Columns total, No % Change)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 14 ' ECM ID
    ws.Columns("C").ColumnWidth = 18 ' Alert ID
    ws.Columns("D").ColumnWidth = 32 ' Old File Name
    ws.Columns("E").ColumnWidth = 32 ' New File Name
    ws.Columns("F").ColumnWidth = 24 ' Old Period
    ws.Columns("G").ColumnWidth = 24 ' New Period
    ws.Columns("H").ColumnWidth = 11 ' Old Count
    ws.Columns("I").ColumnWidth = 11 ' New Count
    ws.Columns("J").ColumnWidth = 12 ' Count Delta
    ws.Columns("K").ColumnWidth = 12 ' Old Alerted
    ws.Columns("L").ColumnWidth = 12 ' New Alerted
    ws.Columns("M").ColumnWidth = 12 ' Alerted Delta
    ws.Columns("N").ColumnWidth = 16 ' New CPs (Pivot)
    ws.Columns("O").ColumnWidth = 38 ' Status Badge
    
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
        .Value = "Automated reconciliation when lookback dates expand | Validates transaction counts, newly alerted transactions, and newly added counterparties in Pivot"
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Color = RGB(100, 116, 139)
        .Interior.Color = RGB(241, 245, 249)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Build 5 Balanced KPI Tiles (Columns B to O)
    CreateKPITile ws, "B7:C9", "ALERTS EVALUATED", "0", "Total compared alert pairs"
    CreateKPITile ws, "D7:F9", "ALERTS WITH COUNT INCREASE", "0", "Lookback count increased [WARN]"
    CreateKPITile ws, "G7:I9", "TOTAL EXTRA TXNS ADDED", "0", "Net transaction volume growth"
    CreateKPITile ws, "J7:L9", "NEW ALERTED TXNS", "0", "Newly captured in-scope [ALERT]"
    CreateKPITile ws, "M7:O9", "NEW COUNTERPARTIES (PIVOT)", "0", "New entities detected in Pivot"
    
    ' Action Buttons (Row 12 across Columns B to O)
    CreateActionButton ws, 12, "B", "D", "[RUN] Run Batch Comparison", "Run_Batch_Comparison", COLOR_ACCENT_BLUE
    CreateActionButton ws, 12, "E", "G", "[OPEN] Compare Single Pair", "Run_Single_Pair_Comparison", RGB(79, 70, 229)
    CreateActionButton ws, 12, "H", "J", "[EXPORT] Export to New Workbook", "Export_To_New_Workbook", RGB(16, 149, 193)
    CreateActionButton ws, 12, "K", "M", "[VIEW] View Discrepancies", "View_Added_Transactions", RGB(13, 148, 136)
    CreateActionButton ws, 12, "N", "O", "[RESET] Reset", "Clear_Comparison_Dashboard", RGB(100, 116, 139)
    
    ' Table Header Row (Columns B to O - 14 Columns total)
    Dim headers As Variant
    headers = Array("ECM ID", "Alert ID", "Old File Name", "New File Name", "Old Date Period", "New Date Period", _
                    "Old Count", "New Count", "Count Delta", "Old Alerted", "New Alerted", "Alerted Delta", "New CPs (Pivot)", "Reconciliation Status")
                    
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
                           ByVal addedTxns As Long, ByVal newAlerted As Long, ByVal newPivotCPs As Long)
    ws.Range("B8").Value = totalAlerts
    ws.Range("D8").Value = increasedCount
    If increasedCount > 0 Then ws.Range("D8").Font.Color = COLOR_AMBER
    
    ws.Range("G8").Value = addedTxns
    If addedTxns > 0 Then ws.Range("G8").Font.Color = COLOR_ACCENT_BLUE
    
    ws.Range("J8").Value = newAlerted
    If newAlerted > 0 Then ws.Range("J8").Font.Color = COLOR_ALERT_RED
    
    ws.Range("M8").Value = newPivotCPs
    If newPivotCPs > 0 Then ws.Range("M8").Font.Color = COLOR_AMBER
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
        
        ws.Range(ws.Cells(r, 11), ws.Cells(r, 13)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(r, 11), ws.Cells(r, 13)).NumberFormat = "#,##0"
        
        ws.Cells(r, 14).HorizontalAlignment = xlCenter ' New CPs (Pivot)
        
        If ws.Cells(r, 10).Value > 0 Then
            ws.Cells(r, 10).Font.Bold = True
            ws.Cells(r, 10).Font.Color = COLOR_ALERT_RED
            ws.Cells(r, 10).Interior.Color = RGB(254, 242, 242)
        End If
        
        If ws.Cells(r, 13).Value > 0 Then
            ws.Cells(r, 13).Font.Bold = True
            ws.Cells(r, 13).Font.Color = COLOR_ALERT_RED
        End If
        
        If InStr(CStr(ws.Cells(r, 14).Value), "New CP") > 0 Then
            ws.Cells(r, 14).Font.Bold = True
            ws.Cells(r, 14).Font.Color = COLOR_AMBER
            ws.Cells(r, 14).Interior.Color = RGB(254, 243, 199)
        End If
        
        With ws.Cells(r, 15)
            .HorizontalAlignment = xlLeft
            .Font.Bold = True
            If InStr(.Value, "NEW ALERTED") > 0 Then
                .Font.Color = COLOR_ALERT_RED
                .Interior.Color = RGB(254, 242, 242)
            ElseIf InStr(.Value, "NEW CP ADDED") > 0 Then
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

Private Sub FormatAuditSheetHeaders(ByVal ws As Worksheet, ByVal bannerTitle As String)
    ws.Cells.Clear
    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo 0
    
    ws.Columns("A").ColumnWidth = 3
    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 10
    
    ws.Range("B1:Z1").Merge
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
End Sub

' =========================================================================
' [HELPER] 14. CROSS-PLATFORM FILE & STRING HELPERS
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
' [DICT] 15. UNIVERSAL KEY-VALUE LOOKUP
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
