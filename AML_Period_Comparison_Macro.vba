Attribute VB_Name = "AML_Period_Comparison"
' =========================================================================
' [AML] AML TRANSACTION MONITORING: LOOKBACK PERIOD BATCH COMPARISON COCKPIT
' =========================================================================
' Version: 2.6 (User Baseline + ECM ID Option + New Counterparty + Export Engine)
'
' Purpose:
'  1. Compares Old vs. New Transaction Files when review lookback dates expand
'     or change to verify:
'       - Did the overall transaction count increase, decrease, or stay the same?
'       - Did the number of ALERTED transactions increase?
'       - What specific transactions were added or dropped?
'       - Did financial volumes/amounts change?
'  2. Provides an interactive, executive-grade Dashboard in Excel with KPI cards,
'     one-click buttons, and conditional formatting.
'  3. Supports BOTH Single-Alert pair selection AND Batch Folder comparison
'     across multiple alerts.
'  4. Mac & Windows cross-platform compatibility (handles folder/file picking
'     and key-value lookups seamlessly on both operating systems).
' =========================================================================
Option Explicit

' --- Design System Palette (Tailored Slate Theme) ---
Private Const COLOR_CANVAS As Long = 16579320       ' Slate 50  RGB(248, 250, 252)
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
           "You can now click the buttons on the dashboard or run 'Run_Batch_Comparison'.", vbInformation, "Dashboard Ready"
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
    
    ' Step 1: Select Folder
    selectedFolder = Pick_Folder("Select folder containing 'Old' and 'New' subfolders (or select the 'Old' folder)")
    If selectedFolder = "" Then Exit Sub
    
    If Right(selectedFolder, 1) <> pSep Then selectedFolder = selectedFolder & pSep
    
    ' Smart Auto-Detection: Check if selected folder contains 'Old' and 'New' subfolders
    If Dir(selectedFolder & "Old", vbDirectory) <> "" And Dir(selectedFolder & "New", vbDirectory) <> "" Then
        oldFolder = selectedFolder & "Old"
        newFolder = selectedFolder & "New"
    ElseIf InStr(1, selectedFolder, pSep & "Old" & pSep, vbTextCompare) > 0 Or Right(selectedFolder, 4) = "Old" & pSep Then
        oldFolder = selectedFolder
        MsgBox "Please select the folder containing the REVISED / NEW transaction files.", vbInformation, "Step 2: Select New Files Folder"
        newFolder = Pick_Folder("Select Folder with NEW Transaction Files")
        If newFolder = "" Then Exit Sub
    Else
        oldFolder = selectedFolder
        MsgBox "Please select the folder containing the REVISED / NEW transaction files.", vbInformation, "Step 2: Select New Files Folder"
        newFolder = Pick_Folder("Select Folder with NEW Transaction Files")
        If newFolder = "" Then Exit Sub
    End If
    
    Execute_Batch_Comparison wbDashboard, oldFolder, newFolder
End Sub

' =========================================================================
' [OPEN] 2. USER-FACING MACRO: RUN SINGLE ALERT COMPARISON
' =========================================================================
Public Sub Run_Single_Pair_Comparison()
    Dim oldFile As String, newFile As String
    Dim wbDashboard As Workbook
    Dim wsDash As Worksheet
    
    Set wbDashboard = ActiveWorkbook
    Set wsDash = Setup_Comparison_Dashboard(wbDashboard)
    
    MsgBox "Please select the OLD / BASELINE transaction file.", vbInformation, "Step 1 of 2: Select Old File"
    oldFile = Pick_Excel_File("Select OLD Transaction File")
    If oldFile = "" Then Exit Sub
    
    MsgBox "Please select the REVISED / NEW lookback transaction file.", vbInformation, "Step 2 of 2: Select New File"
    newFile = Pick_Excel_File("Select NEW Transaction File")
    If newFile = "" Then Exit Sub
    
    Execute_Single_Comparison wbDashboard, oldFile, newFile
End Sub

' =========================================================================
' [RESET] 3. USER-FACING MACRO: RESET / CLEAR DASHBOARD
' =========================================================================
Public Sub Clear_Comparison_Dashboard()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ans As VbMsgBoxResult
    
    ans = MsgBox("Are you sure you want to clear the Comparison Dashboard and all discrepancy records?", vbQuestion + vbYesNo, "Confirm Reset")
    If ans <> vbYes Then Exit Sub
    
    Set wb = ActiveWorkbook
    Setup_Comparison_Dashboard wb, True
    
    ' Remove audit sheets if present
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(SHEET_ADDED_TXNS).Delete
    wb.Worksheets(SHEET_DROPPED_TXNS).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    MsgBox "Dashboard has been reset.", vbInformation, "Cleared"
End Sub

' =========================================================================
' [EXPORT] 3b. USER-FACING MACRO: EXPORT TO STANDALONE FORMATTED WORKBOOK
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
    defaultFileName = "AML_Lookback_Reconciliation_Report_" & timeStamp & ".xlsx"
    
#If Mac Then
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName)
#Else
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, _
                                            FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
                                            Title:="Export AML Comparison Report to New Workbook")
#End If
    If VarType(savePath) = vbBoolean And savePath = False Then Exit Sub
    If finalPath = "" Or finalPath = "False" Then Exit Sub
    
    Dim finalPath As String
    finalPath = finalPath
    If LCase(Right(finalPath, 5)) <> ".xlsx" Then finalPath = finalPath & ".xlsx" 
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Copy Dashboard to new workbook
    wsDash.Copy
    Set wbNew = ActiveWorkbook
    
    ' Strip action button shapes from exported dashboard
    On Error Resume Next
    For Each shp In wbNew.Worksheets(1).Shapes
        If Left(shp.Name, 4) = "Btn_" Then shp.Delete
    Next shp
    On Error GoTo 0
    
    ' Copy Discrepancy Added Txns if present
    On Error Resume Next
    Set wsAdded = wbSource.Worksheets(SHEET_ADDED_TXNS)
    On Error GoTo 0
    If Not wsAdded Is Nothing Then
        wsAdded.Copy After:=wbNew.Worksheets(wbNew.Worksheets.Count)
    End If
    
    ' Copy Discrepancy Dropped Txns if present
    On Error Resume Next
    Set wsDropped = wbSource.Worksheets(SHEET_DROPPED_TXNS)
    On Error GoTo 0
    If Not wsDropped Is Nothing Then
        wsDropped.Copy After:=wbNew.Worksheets(wbNew.Worksheets.Count)
    End If
    
    wbNew.Worksheets(1).Activate
    wbNew.SaveAs Filename:=finalPath, FileFormat:=xlOpenXMLWorkbook
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "AML Lookback Report successfully exported to:" & vbCrLf & vbCrLf & _
           finalPath, vbInformation, "Export Successful"
End Sub


' =========================================================================
' [VIEW] 4. USER-FACING MACRO: VIEW GRANULAR DISCREPANCIES
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
' [ENGINE] 5. CORE BATCH COMPARISON ENGINE
' =========================================================================
Private Sub Execute_Batch_Comparison(ByVal wbDash As Workbook, ByVal oldDir As String, ByVal newDir As String)
    Dim wsDash As Worksheet
    Dim wsAdded As Worksheet, wsDropped As Worksheet
    Dim oldFiles As Collection, newFiles As Collection
    Dim oldMap As Object, newMap As Object
    Dim matchKeys As Collection
    Dim i As Long, curRow As Long
    Dim ecmID As String, alertID As String, matchKey As String, dummyKey As String
    Dim oldFilePath As String, newFilePath As String
    Dim totalAlerts As Long, countIncreasedAlerts As Long
    Dim totalAddedTxns As Long, totalDroppedTxns As Long, totalAlertedAdded As Long
    Dim netVolumeDelta As Double
    Dim stepName As String
    
    On Error GoTo ErrorHandler
    
    stepName = "Initializing Environment"
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    stepName = "Setting Up Worksheets"
    Set wsDash = wbDash.Worksheets(SHEET_DASHBOARD)
    Set wsAdded = GetOrCreateWorksheet(wbDash, SHEET_ADDED_TXNS)
    Set wsDropped = GetOrCreateWorksheet(wbDash, SHEET_DROPPED_TXNS)
    
    FormatAuditSheetHeaders wsAdded, "NEWLY ADDED TRANSACTIONS (EXPANDED LOOKBACK PERIOD)", True
    FormatAuditSheetHeaders wsDropped, "DROPPED / EXCLUDED TRANSACTIONS (REVISED LOOKBACK PERIOD)", False
    
    stepName = "Scanning Folders for Files"
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
    
    stepName = "Indexing and Matching Files"
    Set oldMap = CreateLookupDict()
    Set newMap = CreateLookupDict()
    Set matchKeys = New Collection
    
    Dim fPath As Variant
    For Each fPath In oldFiles
        ExtractFileIDs CStr(fPath), ecmID, alertID, matchKey
        If Not DictExists(oldMap, matchKey) Then
            DictAdd oldMap, matchKey, CStr(fPath)
            matchKeys.Add matchKey
        End If
    Next fPath
    
    For Each fPath In newFiles
        ExtractFileIDs CStr(fPath), ecmID, alertID, matchKey
        If Not DictExists(newMap, matchKey) Then
            DictAdd newMap, matchKey, CStr(fPath)
            If Not DictExists(oldMap, matchKey) Then
                matchKeys.Add matchKey
            End If
        End If
    Next fPath
    
    stepName = "Clearing Old Dashboard Rows"
    curRow = 15
    If wsDash.Cells(curRow, 2).Value <> "" Then
        wsDash.Range("B15:R" & wsDash.Cells(wsDash.Rows.Count, 2).End(xlUp).Row + 5).ClearContents
        wsDash.Range("B15:R" & wsDash.Cells(wsDash.Rows.Count, 2).End(xlUp).Row + 5).ClearFormats
    End If
    
    totalAlerts = 0
    countIncreasedAlerts = 0
    totalAddedTxns = 0
    totalDroppedTxns = 0
    totalAlertedAdded = 0
    netVolumeDelta = 0#
    
    stepName = "Comparing Alert Pairs"
    Dim mKeyVar As Variant
    For Each mKeyVar In matchKeys
        matchKey = CStr(mKeyVar)
        oldFilePath = ""
        newFilePath = ""
        If DictExists(oldMap, matchKey) Then oldFilePath = DictGet(oldMap, matchKey)
        If DictExists(newMap, matchKey) Then newFilePath = DictGet(newMap, matchKey)
        
        If oldFilePath <> "" Then
            ExtractFileIDs oldFilePath, ecmID, alertID, dummyKey
        Else
            ExtractFileIDs newFilePath, ecmID, alertID, dummyKey
        End If
        
        If oldFilePath <> "" And newFilePath <> "" Then
            totalAlerts = totalAlerts + 1
            ProcessPair wsDash, wsAdded, wsDropped, curRow, ecmID, alertID, oldFilePath, newFilePath, _
                        countIncreasedAlerts, totalAddedTxns, totalDroppedTxns, totalAlertedAdded, netVolumeDelta
            curRow = curRow + 1
        ElseIf oldFilePath <> "" Then
            totalAlerts = totalAlerts + 1
            WriteUnmatchedRow wsDash, curRow, ecmID, alertID, GetFileName(oldFilePath), "MISSING IN NEW FOLDER", "UNMATCHED (Missing New File) [WARN]"
            curRow = curRow + 1
        ElseIf newFilePath <> "" Then
            totalAlerts = totalAlerts + 1
            WriteUnmatchedRow wsDash, curRow, ecmID, alertID, "MISSING IN OLD FOLDER", GetFileName(newFilePath), "UNMATCHED (Missing Old File) [WARN]"
            curRow = curRow + 1
        End If
    Next mKeyVar
    
    stepName = "Formatting Dashboard Table"
    FormatBatchTable wsDash, 15, curRow - 1
    
    stepName = "Updating KPI Cards"
    UpdateKPICards wsDash, totalAlerts, countIncreasedAlerts, totalAddedTxns, totalAlertedAdded, netVolumeDelta
    
    wsDash.Activate
    
    MsgBox "Batch Comparison Complete!" & vbCrLf & vbCrLf & _
           "- Total Alerts Processed: " & totalAlerts & vbCrLf & _
           "- Alerts With Increased Count: " & countIncreasedAlerts & vbCrLf & _
           "- Total Transactions Added: " & totalAddedTxns & vbCrLf & _
           "- New Alerted Transactions: " & totalAlertedAdded & vbCrLf & _
           "- Net Financial Volume Delta: " & Format(netVolumeDelta, "$#,##0.00"), vbInformation, "Reconciliation Finished"

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error during Batch Comparison (" & stepName & "): " & vbCrLf & Err.Description, vbCritical, "Error " & Err.Number
    Resume CleanExit
End Sub

' =========================================================================
' [ENGINE] 6. SINGLE PAIR COMPARISON EXECUTION
' =========================================================================
Private Sub Execute_Single_Comparison(ByVal wbDash As Workbook, ByVal oldFilePath As String, ByVal newFilePath As String)
    Dim wsDash As Worksheet
    Dim wsAdded As Worksheet, wsDropped As Worksheet
    Dim ecmID As String, alertID As String, mKey As String
    Dim countIncreasedAlerts As Long, totalAddedTxns As Long, totalDroppedTxns As Long, totalAlertedAdded As Long
    Dim netVolumeDelta As Double
    Dim curRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wsDash = wbDash.Worksheets(SHEET_DASHBOARD)
    Set wsAdded = GetOrCreateWorksheet(wbDash, SHEET_ADDED_TXNS)
    Set wsDropped = GetOrCreateWorksheet(wbDash, SHEET_DROPPED_TXNS)
    
    FormatAuditSheetHeaders wsAdded, "NEWLY ADDED TRANSACTIONS (EXPANDED LOOKBACK PERIOD)", True
    FormatAuditSheetHeaders wsDropped, "DROPPED / EXCLUDED TRANSACTIONS (REVISED LOOKBACK PERIOD)", False
    
    ExtractFileIDs newFilePath, ecmID, alertID, mKey
    curRow = 15
    
    ' Clear table
    wsDash.Range("B15:R30").ClearContents
    wsDash.Range("B15:R30").ClearFormats
    
    ProcessPair wsDash, wsAdded, wsDropped, curRow, ecmID, alertID, oldFilePath, newFilePath, _
                countIncreasedAlerts, totalAddedTxns, totalDroppedTxns, totalAlertedAdded, netVolumeDelta
                
    FormatBatchTable wsDash, 15, 15
    UpdateKPICards wsDash, 1, countIncreasedAlerts, totalAddedTxns, totalAlertedAdded, netVolumeDelta
    
    wsDash.Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Single Alert Comparison Complete!" & vbCrLf & vbCrLf & _
           "ECM ID: " & ecmID & " | Alert ID: " & alertID & vbCrLf & _
           "- Net Count Delta: " & IIf(totalAddedTxns - totalDroppedTxns >= 0, "+", "") & (totalAddedTxns - totalDroppedTxns) & vbCrLf & _
           "- Newly Added Alerted Transactions: " & totalAlertedAdded & vbCrLf & _
           "- Volume Delta: " & Format(netVolumeDelta, "$#,##0.00"), vbInformation, "Reconciliation Finished"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error comparing files: " & Err.Description, vbCritical, "Error"
End Sub

' =========================================================================
' [SEARCH] 7. RECONCILE WORKBOOK PAIR WORKER
' =========================================================================
Private Sub ProcessPair(ByVal wsDash As Worksheet, ByVal wsAdded As Worksheet, ByVal wsDropped As Worksheet, _
                        ByVal rowIdx As Long, ByVal ecmID As String, ByVal alertID As String, _
                        ByVal oldFilePath As String, ByVal newFilePath As String, _
                        ByRef countIncreasedAlerts As Long, ByRef totalAddedTxns As Long, _
                        ByRef totalDroppedTxns As Long, ByRef totalAlertedAdded As Long, _
                        ByRef netVolumeDelta As Double)
                        
    Dim wbOld As Workbook, wbNew As Workbook
    Dim wsOld As Worksheet, wsNew As Worksheet
    Dim oldMap As Object, newMap As Object, oldCPDict As Object
    Dim wasOldOpen As Boolean, wasNewOpen As Boolean
    Dim oldTxCount As Long, newTxCount As Long, countDelta As Long
    Dim oldAlertCount As Long, newAlertCount As Long, alertDelta As Long
    Dim oldAmount As Double, newAmount As Double, amountDelta As Double
    Dim oldMinDate As Variant, oldMaxDate As Variant
    Dim newMinDate As Variant, newMaxDate As Variant
    Dim statusStr As String
    Dim pctChange As Double
    
    ' Open Old File safely
    Set wbOld = GetOrOpenWorkbook(oldFilePath, wasOldOpen)
    Set wsOld = FindTransactionSheet(wbOld)
    Set oldMap = ExtractTransactions(wsOld, oldTxCount, oldAlertCount, oldAmount, oldMinDate, oldMaxDate)
    
    ' Extract Old Counterparties (from Pivot and Txns)
    Set oldCPDict = ExtractOldCounterparties(wbOld, oldMap)
    
    ' Open New File safely
    Set wbNew = GetOrOpenWorkbook(newFilePath, wasNewOpen)
    Set wsNew = FindTransactionSheet(wbNew)
    Set newMap = ExtractTransactions(wsNew, newTxCount, newAlertCount, newAmount, newMinDate, newMaxDate)
    
    ' Close workbooks safely if we opened them
    If Not wasOldOpen Then wbOld.Close SaveChanges:=False
    If Not wasNewOpen Then wbNew.Close SaveChanges:=False
    
    countDelta = newTxCount - oldTxCount
    alertDelta = newAlertCount - oldAlertCount
    amountDelta = newAmount - oldAmount
    pctChange = 0#
    If oldTxCount > 0 Then pctChange = (countDelta / oldTxCount) * 100#
    
    If countDelta > 0 Then
        countIncreasedAlerts = countIncreasedAlerts + 1
        If alertDelta > 0 Then
            statusStr = "COUNT INCREASED (+" & countDelta & ") | NEW ALERTED TXNS (+" & alertDelta & ") [ALERT]"
        Else
            statusStr = "COUNT INCREASED (+" & countDelta & ") [WARN]"
        End If
    ElseIf countDelta < 0 Then
        statusStr = "COUNT DECREASED (" & countDelta & ") [-]"
    Else
        statusStr = "NO COUNT CHANGE (0) [OK]"
    End If
    
    ' Write into Dashboard row
    With wsDash
        .Cells(rowIdx, 2).Value = ecmID
        .Cells(rowIdx, 3).Value = alertID
        .Cells(rowIdx, 4).Value = GetFileName(oldFilePath)
        .Cells(rowIdx, 5).Value = GetFileName(newFilePath)
        .Cells(rowIdx, 6).Value = FormatDateRange(oldMinDate, oldMaxDate, oldFilePath)
        .Cells(rowIdx, 7).Value = FormatDateRange(newMinDate, newMaxDate, newFilePath)
        .Cells(rowIdx, 8).Value = oldTxCount
        .Cells(rowIdx, 9).Value = newTxCount
        .Cells(rowIdx, 10).Value = countDelta
        .Cells(rowIdx, 11).Value = pctChange / 100#
        .Cells(rowIdx, 12).Value = oldAlertCount
        .Cells(rowIdx, 13).Value = newAlertCount
        .Cells(rowIdx, 14).Value = alertDelta
        .Cells(rowIdx, 15).Value = oldAmount
        .Cells(rowIdx, 16).Value = newAmount
        .Cells(rowIdx, 17).Value = amountDelta
        .Cells(rowIdx, 18).Value = statusStr
    End With
    
    ' Record newly added transactions (with Counterparty comparison)
    Dim pairAddedCount As Long, pairAlertedAdded As Long, pairNewCPCount As Long
    RecordAddedTransactions wsAdded, ecmID, alertID, newMap, oldMap, oldCPDict, pairAddedCount, pairAlertedAdded, pairNewCPCount
    
    If pairNewCPCount > 0 Then
        statusStr = statusStr & " | NEW CP (+" & pairNewCPCount & ") [ALERT]"
        wsDash.Cells(rowIdx, 18).Value = statusStr
    End If
    
    ' Record dropped transactions
    Dim pairDroppedCount As Long
    RecordDroppedTransactions wsDropped, ecmID, alertID, oldMap, newMap, pairDroppedCount
    
    totalAddedTxns = totalAddedTxns + pairAddedCount
    totalDroppedTxns = totalDroppedTxns + pairDroppedCount
    totalAlertedAdded = totalAlertedAdded + pairAlertedAdded
    netVolumeDelta = netVolumeDelta + amountDelta
End Sub

' =========================================================================
' [METRICS] 8. EXTRACT TRANSACTIONS FROM SHEET
' =========================================================================
Private Function ExtractTransactions(ByVal ws As Worksheet, ByRef txCount As Long, ByRef alertCount As Long, _
                                     ByRef totalAmt As Double, ByRef minDate As Variant, ByRef maxDate As Variant) As Object
    Dim dict As Object
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim txIdCol As Long, isAlertCol As Long, dateCol As Long, amtCol As Long, accCol As Long, descCol As Long, cpCol As Long
    Dim hText As String
    Dim txID As String, isAlert As String, dtVal As Variant, amtVal As Double
    Dim accNo As String, descText As String, cpName As String
    Dim curDate As Date
    Dim recArray As Variant
    
    Set dict = CreateLookupDict()
    txCount = 0
    alertCount = 0
    totalAmt = 0#
    minDate = Empty
    maxDate = Empty
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Set ExtractTransactions = dict
        Exit Function
    End If
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Auto-detect headers
    txIdCol = 1: isAlertCol = 0: dateCol = 0: amtCol = 0: accCol = 0: descCol = 0: cpCol = 0
    For c = 1 To lastCol
        hText = LCase(Trim(CStr(ws.Cells(1, c).Value)))
        If InStr(hText, "transaction id") > 0 Or InStr(hText, "txn id") > 0 Then
            txIdCol = c
        ElseIf InStr(hText, "is alerted") > 0 Or hText = "alerted" Then
            isAlertCol = c
        ElseIf InStr(hText, "date") > 0 And dateCol = 0 Then
            dateCol = c
        ElseIf (InStr(hText, "amount") > 0 Or InStr(hText, "value") > 0) And amtCol = 0 Then
            amtCol = c
        ElseIf InStr(hText, "account") > 0 And accCol = 0 Then
            accCol = c
        ElseIf (InStr(hText, "description") > 0 Or InStr(hText, "desc") > 0) And descCol = 0 Then
            descCol = c
        ElseIf InStr(hText, "counterparty") > 0 And cpCol = 0 Then
            cpCol = c
        End If
    Next c
    
    ' Fallback defaults based on standard AML layout
    If dateCol = 0 Then dateCol = 9
    If amtCol = 0 Then amtCol = 8
    If isAlertCol = 0 Then isAlertCol = 2
    
    For r = 2 To lastRow
        txID = Trim(CStr(ws.Cells(r, txIdCol).Value))
        If txID <> "" Then
            txCount = txCount + 1
            
            ' Alert status
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
            
            ' Date
            dtVal = ws.Cells(r, dateCol).Value
            If IsDate(dtVal) Then
                curDate = CDate(dtVal)
                If IsEmpty(minDate) Or curDate < minDate Then minDate = curDate
                If IsEmpty(maxDate) Or curDate > maxDate Then maxDate = curDate
            ElseIf IsNumeric(dtVal) And dtVal > 30000 Then
                curDate = CDate(dtVal)
                If IsEmpty(minDate) Or curDate < minDate Then minDate = curDate
                If IsEmpty(maxDate) Or curDate > maxDate Then maxDate = curDate
            End If
            
            ' Amount
            amtVal = 0#
            If amtCol > 0 Then
                If IsNumeric(ws.Cells(r, amtCol).Value) Then
                    amtVal = CDbl(ws.Cells(r, amtCol).Value)
                End If
            End If
            totalAmt = totalAmt + amtVal
            
            ' Other details
            accNo = "": descText = "": cpName = ""
            If accCol > 0 Then accNo = CStr(ws.Cells(r, accCol).Value)
            If descCol > 0 Then descText = CStr(ws.Cells(r, descCol).Value)
            If cpCol > 0 Then cpName = CStr(ws.Cells(r, cpCol).Value)
            
            ' Store in dictionary: Array(TxID, IsAlert, Date, Amount, AccNo, Desc, CP)
            ReDim recArray(0 To 6)
            recArray(0) = txID
            recArray(1) = isAlert
            recArray(2) = dtVal
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
' [AUDIT] 9. AUDIT DETAIL WRITERS
' =========================================================================
Private Sub RecordAddedTransactions(ByVal wsAdded As Worksheet, ByVal ecmID As String, ByVal alertID As String, _
                                    ByVal newMap As Object, ByVal oldMap As Object, _
                                    ByVal oldCPDict As Object, _
                                    ByRef addedCount As Long, ByRef alertedAdded As Long, _
                                    ByRef newCPCount As Long)
    Dim txKeys As Variant
    Dim i As Long, curRow As Long
    Dim txID As String, cpName As String
    Dim rec As Variant
    Dim isNewCP As Boolean
    
    addedCount = 0
    alertedAdded = 0
    newCPCount = 0
    curRow = wsAdded.Cells(wsAdded.Rows.Count, 2).End(xlUp).Row + 1
    If curRow < 4 Then curRow = 4
    
    txKeys = DictKeys(newMap)
    For i = LBound(txKeys) To UBound(txKeys)
        txID = CStr(txKeys(i))
        If Not DictExists(oldMap, txID) Then
            rec = DictGet(newMap, txID)
            addedCount = addedCount + 1
            If CStr(rec(1)) = "Yes" Then alertedAdded = alertedAdded + 1
            
            cpName = Trim(CStr(rec(6)))
            isNewCP = False
            If cpName <> "" And Not oldCPDict Is Nothing Then
                If Not DictExists(oldCPDict, UCase(cpName)) Then
                    isNewCP = True
                    newCPCount = newCPCount + 1
                End If
            End If
            
            wsAdded.Cells(curRow, 2).Value = ecmID
            wsAdded.Cells(curRow, 3).Value = alertID
            wsAdded.Cells(curRow, 4).Value = rec(0) ' Txn ID
            wsAdded.Cells(curRow, 5).Value = rec(1) ' Alerted?
            wsAdded.Cells(curRow, 6).Value = rec(2) ' Date
            wsAdded.Cells(curRow, 7).Value = rec(3) ' Amount
            wsAdded.Cells(curRow, 8).Value = rec(4) ' Account No
            wsAdded.Cells(curRow, 9).Value = rec(5) ' Description
            wsAdded.Cells(curRow, 10).Value = rec(6) ' Counterparty
            wsAdded.Cells(curRow, 11).Value = IIf(isNewCP, "Yes [NEW CP]", "No")
            
            ' Number formatting
            If IsDate(rec(2)) Then wsAdded.Cells(curRow, 6).NumberFormat = "yyyy-mm-dd"
            wsAdded.Cells(curRow, 7).NumberFormat = "$#,##0.00"
            
            ' Alignments
            wsAdded.Cells(curRow, 2).HorizontalAlignment = xlCenter
            wsAdded.Cells(curRow, 3).HorizontalAlignment = xlLeft
            wsAdded.Cells(curRow, 4).HorizontalAlignment = xlCenter
            wsAdded.Cells(curRow, 5).HorizontalAlignment = xlCenter
            wsAdded.Cells(curRow, 6).HorizontalAlignment = xlCenter
            wsAdded.Cells(curRow, 7).HorizontalAlignment = xlRight
            wsAdded.Cells(curRow, 8).HorizontalAlignment = xlLeft
            wsAdded.Cells(curRow, 9).HorizontalAlignment = xlLeft
            wsAdded.Cells(curRow, 10).HorizontalAlignment = xlLeft
            wsAdded.Cells(curRow, 11).HorizontalAlignment = xlCenter
            
            ' Conditional badge for Alerted
            If CStr(rec(1)) = "Yes" Then
                wsAdded.Cells(curRow, 5).Font.Bold = True
                wsAdded.Cells(curRow, 5).Font.Color = COLOR_ALERT_RED
                wsAdded.Cells(curRow, 5).Interior.Color = RGB(254, 242, 242)
            Else
                wsAdded.Cells(curRow, 5).Font.Color = RGB(71, 85, 105)
            End If
            
            ' Conditional badge for New Counterparty
            If isNewCP Then
                wsAdded.Cells(curRow, 10).Font.Bold = True
                wsAdded.Cells(curRow, 11).Font.Bold = True
                wsAdded.Cells(curRow, 11).Font.Color = COLOR_ALERT_RED
                wsAdded.Cells(curRow, 11).Interior.Color = RGB(254, 242, 242)
            Else
                wsAdded.Cells(curRow, 11).Font.Color = COLOR_SUCCESS_GREEN
            End If
            
            curRow = curRow + 1
        End If
    Next i
End Sub

Private Sub RecordDroppedTransactions(ByVal wsDropped As Worksheet, ByVal ecmID As String, ByVal alertID As String, _
                                      ByVal oldMap As Object, ByVal newMap As Object, _
                                      ByRef droppedCount As Long)
    Dim txKeys As Variant
    Dim i As Long, curRow As Long
    Dim txID As String
    Dim rec As Variant
    
    droppedCount = 0
    curRow = wsDropped.Cells(wsDropped.Rows.Count, 2).End(xlUp).Row + 1
    If curRow < 4 Then curRow = 4
    
    txKeys = DictKeys(oldMap)
    For i = LBound(txKeys) To UBound(txKeys)
        txID = CStr(txKeys(i))
        If Not DictExists(newMap, txID) Then
            rec = DictGet(oldMap, txID)
            droppedCount = droppedCount + 1
            
            wsDropped.Cells(curRow, 2).Value = ecmID
            wsDropped.Cells(curRow, 3).Value = alertID
            wsDropped.Cells(curRow, 4).Value = rec(0) ' Txn ID
            wsDropped.Cells(curRow, 5).Value = rec(1) ' Alerted?
            wsDropped.Cells(curRow, 6).Value = rec(2) ' Date
            wsDropped.Cells(curRow, 7).Value = rec(3) ' Amount
            wsDropped.Cells(curRow, 8).Value = rec(4) ' Account No
            wsDropped.Cells(curRow, 9).Value = rec(5) ' Description
            wsDropped.Cells(curRow, 10).Value = rec(6) ' Counterparty
            
            If IsDate(rec(2)) Then wsDropped.Cells(curRow, 6).NumberFormat = "yyyy-mm-dd"
            wsDropped.Cells(curRow, 7).NumberFormat = "$#,##0.00"
            
            wsDropped.Cells(curRow, 2).HorizontalAlignment = xlCenter
            wsDropped.Cells(curRow, 3).HorizontalAlignment = xlLeft
            wsDropped.Cells(curRow, 4).HorizontalAlignment = xlCenter
            wsDropped.Cells(curRow, 5).HorizontalAlignment = xlCenter
            wsDropped.Cells(curRow, 6).HorizontalAlignment = xlCenter
            wsDropped.Cells(curRow, 7).HorizontalAlignment = xlRight
            wsDropped.Cells(curRow, 8).HorizontalAlignment = xlLeft
            wsDropped.Cells(curRow, 9).HorizontalAlignment = xlLeft
            wsDropped.Cells(curRow, 10).HorizontalAlignment = xlLeft
            
            curRow = curRow + 1
        End If
    Next i
End Sub

' =========================================================================
' [UI] 10. DASHBOARD FORMATTING & BUTTON SETUP
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
    
    ' Remove old shapes / buttons
    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0
    
    ' Enable gridlines safely on active window
    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo 0
    
    ' Column Widths (Columns B to R: 17 Data Columns)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 14 ' ECM ID
    ws.Columns("C").ColumnWidth = 18 ' Alert ID
    ws.Columns("D").ColumnWidth = 30 ' Old File
    ws.Columns("E").ColumnWidth = 30 ' New File
    ws.Columns("F").ColumnWidth = 24 ' Old Period
    ws.Columns("G").ColumnWidth = 24 ' New Period
    ws.Columns("H").ColumnWidth = 12 ' Old Count
    ws.Columns("I").ColumnWidth = 12 ' New Count
    ws.Columns("J").ColumnWidth = 12 ' Delta Count
    ws.Columns("K").ColumnWidth = 11 ' % Change
    ws.Columns("L").ColumnWidth = 12 ' Old Alerted
    ws.Columns("M").ColumnWidth = 12 ' New Alerted
    ws.Columns("N").ColumnWidth = 12 ' Alerted Delta
    ws.Columns("O").ColumnWidth = 18 ' Old Amount
    ws.Columns("P").ColumnWidth = 18 ' New Amount
    ws.Columns("Q").ColumnWidth = 18 ' Amount Delta
    ws.Columns("R").ColumnWidth = 40 ' Status Badge
    
    ' Row Heights
    ws.Rows("1:3").RowHeight = 12
    ws.Rows(4).RowHeight = 32
    ws.Rows(5).RowHeight = 20
    ws.Rows(6).RowHeight = 10
    ws.Rows("7:10").RowHeight = 22 ' KPI Cards
    ws.Rows(11).RowHeight = 12
    ws.Rows(12).RowHeight = 30     ' Button Bar
    ws.Rows(13).RowHeight = 12
    ws.Rows(14).RowHeight = 26     ' Table Header
    
    ' Master Banner Header
    ws.Range("B4:R4").Merge
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
    
    ws.Range("B5:R5").Merge
    With ws.Range("B5")
        .Value = "Automated discrepancy verification when review period dates expand | Quantifies transaction count increases, newly alerted transactions, and volume shifts"
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Font.Color = RGB(100, 116, 139) ' Slate 500
        .Interior.Color = RGB(241, 245, 249) ' Slate 100
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Build KPI Tiles (Spanning Columns B to R seamlessly)
    CreateKPITile ws, "B7:D9", "ALERTS EVALUATED", "0", "Total compared alert files"
    CreateKPITile ws, "E7:G9", "ALERTS WITH INCREASED COUNT", "0", "Lookback count increased ([WARN])"
    CreateKPITile ws, "H7:J9", "TOTAL TXNS ADDED", "0", "Net transaction volume growth"
    CreateKPITile ws, "K7:M9", "NEW ALERTED TXNS", "0", "Newly captured in-scope alerts ([ALERT])"
    CreateKPITile ws, "N7:R9", "NET VOLUME DELTA ($)", "$0.00", "Total financial difference"
    
    ' Action Buttons (Spanning Columns B to R seamlessly on Row 12)
    CreateActionButton ws, 12, "B", "D", "[RUN] Run Batch Comparison", "Run_Batch_Comparison", COLOR_ACCENT_BLUE
    CreateActionButton ws, 12, "E", "G", "[OPEN] Compare Single Pair", "Run_Single_Pair_Comparison", RGB(79, 70, 229)
    CreateActionButton ws, 12, "H", "J", "[EXPORT] Export to Workbook", "Export_To_New_Workbook", RGB(16, 149, 193)
    CreateActionButton ws, 12, "K", "M", "[VIEW] View Added Txns", "View_Added_Transactions", RGB(13, 148, 136)
    CreateActionButton ws, 12, "N", "R", "[RESET] Reset Dashboard", "Clear_Comparison_Dashboard", RGB(100, 116, 139)
    
    ' Table Header Row
    Dim headers As Variant
    headers = Array("ECM ID", "Alert ID", "Old File Name", "New File Name", "Old Date Period", "New Date Period", _
                    "Old Count", "New Count", "Count Delta", "% Change", "Old Alerted", "New Alerted", "Alerted Delta", _
                    "Old Amount ($)", "New Amount ($)", "Amount Delta ($)", "Reconciliation Status")
                    
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
    
    ' Title Row
    With firstCell
        .Value = title
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Font.Bold = True
        .Font.Color = RGB(100, 116, 139)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Value Row
    With firstCell.Offset(1, 0)
        .Value = valStr
        .Font.Name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(15, 23, 42)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Subtitle Row
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
                           ByVal addedTxns As Long, ByVal newAlerted As Long, ByVal volumeDelta As Double)
    ws.Range("B8").Value = totalAlerts
    ws.Range("E8").Value = increasedCount
    If increasedCount > 0 Then ws.Range("E8").Font.Color = COLOR_AMBER
    
    ws.Range("H8").Value = addedTxns
    If addedTxns > 0 Then ws.Range("H8").Font.Color = COLOR_ACCENT_BLUE
    
    ws.Range("K8").Value = newAlerted
    If newAlerted > 0 Then ws.Range("K8").Font.Color = COLOR_ALERT_RED
    
    ws.Range("N8").Value = Format(volumeDelta, "$#,##0.00")
    If volumeDelta > 0 Then ws.Range("N8").Font.Color = COLOR_SUCCESS_GREEN
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
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Private Sub FormatBatchTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    If endRow < startRow Then Exit Sub
    
    Dim rngTable As Range
    Set rngTable = ws.Range("B" & startRow & ":R" & endRow)
    
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
        
        ' Alignments
        ws.Cells(r, 2).HorizontalAlignment = xlCenter ' ECM ID
        ws.Cells(r, 3).HorizontalAlignment = xlLeft   ' Alert ID
        ws.Cells(r, 4).HorizontalAlignment = xlLeft   ' Old File
        ws.Cells(r, 5).HorizontalAlignment = xlLeft   ' New File
        ws.Cells(r, 6).HorizontalAlignment = xlCenter ' Old Period
        ws.Cells(r, 7).HorizontalAlignment = xlCenter ' New Period
        
        ' Counts (Columns 8, 9, 10: Old, New, Delta)
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 10)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 10)).NumberFormat = "#,##0"
        
        ' % Change (Column 11)
        ws.Cells(r, 11).HorizontalAlignment = xlRight
        ws.Cells(r, 11).NumberFormat = "0.00%"
        
        ' Alerted counts (Columns 12, 13, 14: Old, New, Delta)
        ws.Range(ws.Cells(r, 12), ws.Cells(r, 14)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(r, 12), ws.Cells(r, 14)).NumberFormat = "#,##0"
        
        ' Amounts (Columns 15, 16, 17: Old, New, Delta)
        ws.Range(ws.Cells(r, 15), ws.Cells(r, 17)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(r, 15), ws.Cells(r, 17)).NumberFormat = "$#,##0.00"
        
        ' Delta styling
        If IsNumeric(ws.Cells(r, 10).Value) Then
            If ws.Cells(r, 10).Value > 0 Then
                ws.Cells(r, 10).Font.Bold = True
                ws.Cells(r, 10).Font.Color = COLOR_ALERT_RED
                ws.Cells(r, 10).Interior.Color = RGB(254, 242, 242)
            End If
        End If
        
        If IsNumeric(ws.Cells(r, 14).Value) Then
            If ws.Cells(r, 14).Value > 0 Then
                ws.Cells(r, 14).Font.Bold = True
                ws.Cells(r, 14).Font.Color = COLOR_ALERT_RED
            End If
        End If
        
        ' Status Column Badge (Column 18)
        With ws.Cells(r, 18)
            .HorizontalAlignment = xlLeft
            .Font.Bold = True
            If InStr(.Value, "NEW ALERTED TXNS") > 0 Or InStr(.Value, "NEW CP") > 0 Then
                .Font.Color = COLOR_ALERT_RED
                .Interior.Color = RGB(254, 242, 242)
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
    wsDash.Cells(rowIdx, 15).Value = "-"
    wsDash.Cells(rowIdx, 16).Value = "-"
    wsDash.Cells(rowIdx, 17).Value = "-"
    wsDash.Cells(rowIdx, 18).Value = statusStr
    wsDash.Cells(rowIdx, 18).Font.Bold = True
    wsDash.Cells(rowIdx, 18).Font.Color = COLOR_ALERT_RED
End Sub

Private Sub FormatAuditSheetHeaders(ByVal ws As Worksheet, ByVal bannerTitle As String, Optional ByVal isAddedSheet As Boolean = True)
    ws.Cells.Clear
    ' Enable gridlines safely on active window
    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    On Error GoTo 0
    
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 16 ' ECM ID
    ws.Columns("C").ColumnWidth = 18 ' Alert ID
    ws.Columns("D").ColumnWidth = 16 ' Txn ID
    ws.Columns("E").ColumnWidth = 12 ' Alerted?
    ws.Columns("F").ColumnWidth = 14 ' Date
    ws.Columns("G").ColumnWidth = 16 ' Amount
    ws.Columns("H").ColumnWidth = 24 ' Account No
    ws.Columns("I").ColumnWidth = 28 ' Description
    ws.Columns("J").ColumnWidth = 32 ' Counterparty
    
    Dim lastColLetter As String
    Dim headers As Variant
    
    If isAddedSheet Then
        ws.Columns("K").ColumnWidth = 20 ' New Counterparty?
        lastColLetter = "K"
        headers = Array("ECM ID", "Alert ID", "Transaction ID", "Is Alerted?", "Transaction Date", "Amount ($)", "Account No", "Transaction Description", "Counterparty Name", "New Counterparty?")
    Else
        lastColLetter = "J"
        headers = Array("ECM ID", "Alert ID", "Transaction ID", "Is Alerted?", "Transaction Date", "Amount ($)", "Account No", "Transaction Description", "Counterparty Name")
    End If
    
    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 8
    ws.Rows(3).RowHeight = 24
    
    ws.Range("B1:" & lastColLetter & "1").Merge
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
    
    Dim c As Long
    For c = 0 To UBound(headers)
        With ws.Cells(3, c + 2)
            .Value = headers(c)
            .Font.Name = "Segoe UI"
            .Font.Size = 9.5
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = RGB(30, 41, 59) ' Slate 800
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
End Sub

' =========================================================================
' [HELPER] 11. CROSS-PLATFORM HELPERS (Mac & Windows)
' =========================================================================

' =========================================================================
' [CP] EXTRACT COUNTERPARTIES FROM OLD FILE (PIVOT + TXNS)
' =========================================================================
Private Function ExtractOldCounterparties(ByVal wb As Workbook, ByVal txMap As Object) As Object
    Dim cpDict As Object
    Dim wsPivot As Worksheet
    Dim r As Long, lastRow As Long, cpVal As String
    Dim txKeys As Variant, i As Long, rec As Variant
    
    Set cpDict = CreateLookupDict()
    If wb Is Nothing Then
        Set ExtractOldCounterparties = cpDict
        Exit Function
    End If
    
    ' Strategy 1: Check Pivot sheet if present
    On Error Resume Next
    Set wsPivot = wb.Worksheets("Pivot")
    If wsPivot Is Nothing Then
        Dim s As Worksheet
        For Each s In wb.Worksheets
            If InStr(1, s.Name, "pivot", vbTextCompare) > 0 Then Set wsPivot = s: Exit For
        Next s
    End If
    On Error GoTo 0
    
    If Not wsPivot Is Nothing Then
        On Error Resume Next
        lastRow = wsPivot.Cells(wsPivot.Rows.Count, 1).End(xlUp).Row
        For r = 4 To lastRow
            If Not IsError(wsPivot.Cells(r, 1).Value) Then
                cpVal = Trim(CStr(wsPivot.Cells(r, 1).Value))
                If cpVal <> "" And UCase(cpVal) <> "ROW LABELS" And UCase(cpVal) <> "GRAND TOTAL" And UCase(cpVal) <> "TOTAL" And UCase(cpVal) <> "(BLANK)" And UCase(cpVal) <> "CR" And UCase(cpVal) <> "DR" Then
                    If Not DictExists(cpDict, UCase(cpVal)) Then DictAdd cpDict, UCase(cpVal), cpVal
                End If
            End If
        Next r
        On Error GoTo 0
    End If
    
    ' Strategy 2: Include all counterparties from transactions map
    If Not txMap Is Nothing Then
        txKeys = DictKeys(txMap)
        For i = LBound(txKeys) To UBound(txKeys)
            rec = DictGet(txMap, CStr(txKeys(i)))
            If IsArray(rec) Then
                If UBound(rec) >= 6 Then
                    cpVal = Trim(CStr(rec(6)))
                    If cpVal <> "" Then
                        If Not DictExists(cpDict, UCase(cpVal)) Then DictAdd cpDict, UCase(cpVal), cpVal
                    End If
                End If
            End If
        Next i
    End If
    
    Set ExtractOldCounterparties = cpDict
End Function

Private Function GetOrOpenWorkbook(ByVal fPath As String, ByRef wasAlreadyOpen As Boolean) As Workbook
    Dim wb As Workbook, fName As String
    fName = GetFileName(fPath)
    wasAlreadyOpen = False
    
    On Error Resume Next
    Set wb = Workbooks(fName)
    On Error GoTo 0
    
    If Not wb Is Nothing Then
        wasAlreadyOpen = True
        Set GetOrOpenWorkbook = wb
        Exit Function
    End If
    
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=fPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0
    
    Set GetOrOpenWorkbook = wb
End Function

Private Function Pick_Folder(ByVal promptTitle As String) As String
#If Mac Then
    ' Native AppleScript on macOS
    Dim script As String
    Dim result As String
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
    Dim sFile As String
    Dim pSep As String
    
    pSep = Application.PathSeparator
    If Right(folderPath, 1) <> pSep Then folderPath = folderPath & pSep
    
    sFile = Dir(folderPath & "*.xlsx")
    Do While sFile <> ""
        If Left(sFile, 2) <> "~$" And InStr(1, sFile, "AML_Period_Comparison", vbTextCompare) = 0 Then
            col.Add folderPath & sFile
        End If
        sFile = Dir()
    Loop
    
    sFile = Dir(folderPath & "*.xlsm")
    Do While sFile <> ""
        If Left(sFile, 2) <> "~$" And InStr(1, sFile, "AML_Period_Comparison", vbTextCompare) = 0 Then
            col.Add folderPath & sFile
        End If
        sFile = Dir()
    Loop
    
    Set ListExcelFiles = col
End Function

Private Sub ExtractFileIDs(ByVal fPath As String, ByRef ecmID As String, ByRef alertID As String, ByRef matchKey As String)
    Dim fName As String, baseName As String
    Dim uName As String
    Dim dotPos As Long, alertPos As Long, sepPos As Long, endPos As Long
    Dim prefix As String, remStr As String, ch As String
    Dim i As Long
    
    fName = GetFileName(fPath)
    dotPos = InStrRev(fName, ".")
    If dotPos > 0 Then
        baseName = Left(fName, dotPos - 1)
    Else
        baseName = fName
    End If
    
    uName = UCase(baseName)
    alertPos = InStr(1, uName, "ALERT")
    
    ecmID = "-"
    alertID = ""
    matchKey = ""
    
    If alertPos > 0 Then
        ' Check for ECM ID before ALERT e.g. 138896_ALERT2365101
        If alertPos > 1 Then
            prefix = Left(baseName, alertPos - 1)
            Do While Len(prefix) > 0 And (Right(prefix, 1) = "_" Or Right(prefix, 1) = "-" Or Right(prefix, 1) = " ")
                prefix = Left(prefix, Len(prefix) - 1)
            Loop
            
            If Len(prefix) > 0 Then
                sepPos = InStrRev(prefix, "_")
                If sepPos = 0 Then sepPos = InStrRev(prefix, "-")
                If sepPos = 0 Then sepPos = InStrRev(prefix, " ")
                If sepPos > 0 Then
                    ecmID = Trim(Mid(prefix, sepPos + 1))
                Else
                    ecmID = Trim(prefix)
                End If
            End If
        End If
        
        ' Extract ALERT token e.g. ALERT2365101
        remStr = Mid(baseName, alertPos)
        endPos = 0
        For i = 1 To Len(remStr)
            ch = Mid(remStr, i, 1)
            If ch = "_" Or ch = " " Or ch = "-" Or ch = "(" Or ch = "." Then
                endPos = i - 1
                Exit For
            End If
        Next i
        
        If endPos > 0 Then
            alertID = Trim(Left(remStr, endPos))
        Else
            alertID = Trim(remStr)
        End If
        
        If ecmID <> "-" And ecmID <> "" Then
            matchKey = ecmID & "_" & alertID
        Else
            matchKey = alertID
        End If
    Else
        ' No ALERT found in filename
        sepPos = InStr(baseName, "_")
        If sepPos > 0 Then
            prefix = Left(baseName, sepPos - 1)
            If IsNumeric(prefix) Then
                ecmID = prefix
                alertID = Mid(baseName, sepPos + 1)
                matchKey = baseName
            Else
                ecmID = "-"
                alertID = baseName
                matchKey = baseName
            End If
        Else
            ecmID = "-"
            alertID = baseName
            matchKey = baseName
        End If
    End If
    
    If alertID = "" Then alertID = baseName
    If matchKey = "" Then matchKey = alertID
End Sub

Private Function GetFileName(ByVal fPath As String) As String
    Dim pSep As String, pos As Long
    pSep = Application.PathSeparator
    pos = InStrRev(fPath, pSep)
    If pos > 0 Then
        GetFileName = Mid(fPath, pos + 1)
    Else
        GetFileName = fPath
    End If
End Function

Private Function FindTransactionSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim sName As String
    
    ' Priority 1: Check known names
    For Each ws In wb.Worksheets
        sName = LCase(ws.Name)
        If InStr(sName, "lookback") > 0 Or InStr(sName, "raw transactions") > 0 Or InStr(sName, "inscope") > 0 Then
            Set FindTransactionSheet = ws
            Exit Function
        End If
    Next ws
    
    ' Priority 2: Check for header 'Transaction ID'
    For Each ws In wb.Worksheets
        If ws.Cells(1, 1).Value <> "" Then
            If InStr(LCase(CStr(ws.Cells(1, 1).Value)), "transaction") > 0 Then
                Set FindTransactionSheet = ws
                Exit Function
            End If
        End If
    Next ws
    
    ' Priority 3: First non-pivot sheet
    For Each ws In wb.Worksheets
        sName = LCase(ws.Name)
        If InStr(sName, "pivot") = 0 And InStr(sName, "dashboard") = 0 Then
            Set FindTransactionSheet = ws
            Exit Function
        End If
    Next ws
    
    Set FindTransactionSheet = wb.Worksheets(1)
End Function

Private Function FormatDateRange(ByVal dMin As Variant, ByVal dMax As Variant, Optional ByVal fPath As String = "") As String
    If IsDate(dMin) And IsDate(dMax) Then
        FormatDateRange = Format(dMin, "yyyy-mm-dd") & " to " & Format(dMax, "yyyy-mm-dd")
        Exit Function
    End If
    
    ' Fallback: Extract date range from filename if present (e.g. 06.02.2025 to 06.03.2026)
    If fPath <> "" Then
        Dim fName As String, posTo As Long
        fName = GetFileName(fPath)
        posTo = InStr(1, fName, " to ", vbTextCompare)
        If posTo > 0 Then
            Dim leftPart As String, rightPart As String
            Dim s1 As Long, s2 As Long
            leftPart = Trim(Left(fName, posTo - 1))
            rightPart = Trim(Mid(fName, posTo + 4))
            
            s1 = InStrRev(leftPart, " ")
            If s1 > 0 Then leftPart = Mid(leftPart, s1 + 1)
            leftPart = Replace(leftPart, "(", "")
            
            s2 = InStr(rightPart, " ")
            If s2 = 0 Then s2 = InStr(rightPart, ")")
            If s2 = 0 Then s2 = InStr(rightPart, ".")
            If s2 > 0 Then rightPart = Left(rightPart, s2 - 1)
            
            If leftPart <> "" And rightPart <> "" Then
                FormatDateRange = leftPart & " to " & rightPart & " (file)"
                Exit Function
            End If
        End If
    End If
    
    FormatDateRange = "N/A"
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
' [DICT] 12. UNIVERSAL KEY-VALUE LOOKUP (Zero-crash Mac/Win implementation)
' =========================================================================
Private Function CreateLookupDict() As Object
    On Error Resume Next
    Set CreateLookupDict = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    
    If CreateLookupDict Is Nothing Then
        ' Fallback on Mac: Use VBA Collection
        Set CreateLookupDict = New Collection
    End If
End Function

Private Function DictExists(ByVal d As Object, ByVal k As String) As Boolean
    If TypeName(d) = "Dictionary" Then
        DictExists = d.Exists(k)
    Else
        ' Collection lookup
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

Private Function DictKeys(ByVal d As Object) As Variant
    If TypeName(d) = "Dictionary" Then
        DictKeys = d.Keys
    Else
        ' Collection keys extraction
        Dim keysArr() As String
        Dim count As Long, i As Long
        count = d.Count
        If count = 0 Then
            DictKeys = Array()
            Exit Function
        End If
        ReDim keysArr(0 To count - 1)
        ' For Collections with items as arrays where recArray(0) is key
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
