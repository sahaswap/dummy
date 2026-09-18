Attribute VB_Name = "AML_Period_Comparison"
' =========================================================================
' [AML] AML TRANSACTION MONITORING: LOOKBACK PERIOD BATCH COMPARISON COCKPIT
' =========================================================================
' Version: 3.0
'
' Purpose:
'  1. Compares Old vs. New transaction files when review lookback dates expand
'     or change, and answers:
'       - Did the transaction count increase, decrease, or stay the same?
'       - Did the number of ALERTED transactions change?
'       - Which transactions were added, dropped, or changed in place
'         (alert flag flips, amount / date / direction / counterparty edits)?
'       - Do added transactions involve counterparties never seen before?
'       - Did gross financial volume change?
'  2. Interactive dashboard with KPI cards, one-click buttons and status badges.
'  3. Single-pair and batch-folder comparison.
'  4. Windows and Mac (Excel 2016+) compatible.
'
' Changes in 3.0 (vs 2.6):
'  - Export macro compiles and works (module previously failed to compile).
'  - Old and New files with the SAME file name are compared correctly (the old
'    workbook was silently re-used as the "new" data before).
'  - Alert IDs like "Alert_123", "ALERT-123", "ALERT 123" parse correctly;
'    duplicate match keys are reported instead of silently skipped.
'  - Text dates (e.g. "10/31/2025") parsed explicitly as TEXT_DATE_ORDER,
'    independent of the PC's regional settings. Bad cells never crash a run.
'  - One unreadable file no longer aborts the batch: it gets an ERROR row.
'  - Detects in-place changes to existing transactions (new audit sheet
'    "Discrepancy_Modified_Txns"), duplicate transaction IDs, missing columns,
'    mixed currencies.
'  - Counterparty is read from a "Counterparty" column, or derived from
'    Dr/Cr + Originator/Beneficiary Name when that column is absent. Known
'    counterparties are gathered from every sheet and pivot of the OLD file.
'    Names are compared after normalising case, punctuation and SWIFT line
'    tags ("1/NAME" = "NAME").
'  - Column detection by header text with priorities; no silent hard-coded
'    column fallbacks. Header row may be anywhere in the first 10 rows.
'  - Array-based reads/writes (much faster on large files).
'  - Dashboard fully cleared between runs; KPI cards reset every run.
'  - Buttons work when the macro lives in another workbook (PERSONAL.XLSB).
'
' Settings: see "User Settings" below.
' =========================================================================
Option Explicit

' --- User Settings ---
' How text dates such as "10/07/2025" are read. "MDY" = US (month first),
' "DMY" = day first. Real Excel dates are never affected by this setting.
Private Const TEXT_DATE_ORDER As String = "MDY"
' Amount differences at or below this are treated as equal.
Private Const AMOUNT_TOLERANCE As Double = 0.005
' How many rows from the top of a sheet are searched for the header row.
Private Const HEADER_SCAN_ROWS As Long = 10

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
Private Const COLOR_TEXT_DARK As Long = 2758415     ' Slate 900 RGB(15, 23, 42)
Private Const COLOR_TEXT_MUTED As Long = 6903111    ' Slate 600 RGB(71, 85, 105)
Private Const COLOR_FILL_RED As Long = 15921918     ' RGB(254, 242, 242)
Private Const COLOR_FILL_AMBER As Long = 13104126   ' RGB(254, 243, 199)
Private Const COLOR_FILL_GREEN As Long = 16055792   ' RGB(240, 253, 244)

#If Mac Then
Private Const FONT_UI As String = "Calibri"
#Else
Private Const FONT_UI As String = "Segoe UI"
#End If

' --- Sheet Names ---
Private Const SHEET_DASHBOARD As String = "Comparison Dashboard"
Private Const SHEET_ADDED_TXNS As String = "Discrepancy_Added_Txns"
Private Const SHEET_DROPPED_TXNS As String = "Discrepancy_Dropped_Txns"
Private Const SHEET_MODIFIED_TXNS As String = "Discrepancy_Modified_Txns"

' --- Dashboard Layout ---
Private Const DASH_VERSION As String = "AML-PERIOD-COMPARISON-3.0"
Private Const DASH_TITLE As String = "AML TRANSACTION MONITORING - LOOKBACK PERIOD BATCH RECONCILIATION COCKPIT"
Private Const FIRST_DATA_ROW As Long = 15
Private Const DASH_COLS As Long = 17               ' columns B:R

' --- Audit Sheet Kinds ---
Private Const AUDIT_ADDED As Long = 1
Private Const AUDIT_DROPPED As Long = 2
Private Const AUDIT_MODIFIED As Long = 3

' --- Header Field IDs ---
Private Const FLD_TXID As Long = 1
Private Const FLD_ALERT As Long = 2
Private Const FLD_DATE As Long = 3
Private Const FLD_AMOUNT As Long = 4
Private Const FLD_ACCOUNT As Long = 5
Private Const FLD_DESC As Long = 6
Private Const FLD_CP As Long = 7
Private Const FLD_DRCR As Long = 8
Private Const FLD_CCY As Long = 9
Private Const FLD_ORIG As Long = 10
Private Const FLD_BEN As Long = 11

' --- Transaction Record Layout (0-based array) ---
Private Const F_TXID As Long = 0
Private Const F_ALERT As Long = 1
Private Const F_DATE As Long = 2
Private Const F_AMT As Long = 3
Private Const F_ACC As Long = 4
Private Const F_DESC As Long = 5
Private Const F_CP As Long = 6
Private Const F_DRCR As Long = 7
Private Const F_CCY As Long = 8

' --- Types ---
Private Type ColMap
    TxID As Long
    Alert As Long
    TxDate As Long
    Amount As Long
    Account As Long
    Desc As Long
    CP As Long
    DrCr As Long
    Ccy As Long
    OrigName As Long
    BenName As Long
End Type

Private Type TxnSet
    Map As Object              ' transaction key -> record array
    SheetName As String
    RowCount As Long           ' rows with a transaction ID (incl. duplicates)
    UniqueCount As Long
    DupCount As Long           ' extra rows re-using an existing transaction ID
    AlertCount As Long         ' unique alerted transactions
    TotalAmt As Double         ' gross volume (sum of absolute amounts)
    MinDate As Variant
    MaxDate As Variant
    BadDates As Long           ' non-blank date cells that could not be read
    Currencies As String       ' e.g. "USD" or "EUR, USD"
    MissingCols As String
    CPSource As String
    HasAlert As Boolean
    HasDate As Boolean
    HasAmount As Boolean
    HasAccount As Boolean
    HasDrCr As Boolean
    HasCcy As Boolean
    HasCP As Boolean
End Type

Private Type PairResult
    AddedCount As Long
    AddedAlerted As Long
    DroppedCount As Long
    DroppedAlerted As Long
    FlipToYes As Long
    FlipToNo As Long
    ModifiedTxns As Long       ' existing transactions with any change
    FieldChangedTxns As Long   ' existing transactions with a non-alert-flag change
    NewCPTxns As Long
    NewCPDistinct As Long
    CPUnknown As Long
End Type

Private Type BatchTotals
    Evaluated As Long
    Unmatched As Long
    DupFiles As Long
    Errors As Long
    IncreasedCount As Long
    ChangedPairs As Long
    Added As Long
    Dropped As Long
    AddedAlerted As Long
    FlipToYes As Long
    Modified As Long
    NewCP As Long
    VolumeDelta As Double
    Currencies As String
End Type

Private Type AppState
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    EnableEvents As Boolean
    Calculation As Long
    HasCalculation As Boolean
End Type

' =========================================================================
' [UI] 0. USER-FACING MACRO: CREATE / REBUILD THE DASHBOARD
' =========================================================================
Public Sub Setup_Dashboard()
    Dim wb As Workbook, ws As Worksheet
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open or create a workbook first.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    Set ws = FindSheet(wb, SHEET_DASHBOARD)
    If Not ws Is Nothing Then
        If DashboardHasResults(ws) Then
            If MsgBox("Rebuilding the dashboard clears the current results table." & vbCrLf & vbCrLf & _
                      "Continue?", vbQuestion + vbYesNo, "Rebuild Dashboard") <> vbYes Then Exit Sub
        End If
    End If

    Set ws = Setup_Comparison_Dashboard(wb, True)
    ws.Activate
    MsgBox "Comparison Dashboard has been created successfully!" & vbCrLf & vbCrLf & _
           "Click the buttons on the dashboard or run 'Run_Batch_Comparison'.", vbInformation, "Dashboard Ready"
End Sub

' =========================================================================
' [RUN] 1. USER-FACING MACRO: RUN BATCH COMPARISON ACROSS ALL ALERTS
' =========================================================================
Public Sub Run_Batch_Comparison()
    Dim selectedFolder As String, oldFolder As String, newFolder As String
    Dim pSep As String
    Dim wbDashboard As Workbook
    Dim jobs As Collection

    Set wbDashboard = ActiveWorkbook
    If wbDashboard Is Nothing Then
        MsgBox "Open the workbook that should hold the dashboard first.", vbExclamation, "No Workbook"
        Exit Sub
    End If
    pSep = Application.PathSeparator

    ' Step 1: Select folder (parent with Old/New subfolders, or the Old folder itself)
    selectedFolder = Pick_Folder("Select the folder containing 'Old' and 'New' subfolders (or select the 'Old' folder)")
    If selectedFolder = "" Then Exit Sub
    GrantMacAccess Array(selectedFolder)

    If FolderExists(selectedFolder & "Old") And FolderExists(selectedFolder & "New") Then
        oldFolder = selectedFolder & "Old" & pSep
        newFolder = selectedFolder & "New" & pSep
    Else
        oldFolder = selectedFolder
        MsgBox "Selected OLD / BASELINE folder:" & vbCrLf & oldFolder & vbCrLf & vbCrLf & _
               "Now select the folder containing the REVISED / NEW transaction files.", vbInformation, "Step 2: Select New Files Folder"
        newFolder = Pick_Folder("Select the folder with the NEW transaction files")
        If newFolder = "" Then Exit Sub
    End If

    If StrComp(oldFolder, newFolder, vbTextCompare) = 0 Then
        MsgBox "The Old and New folders are the same folder:" & vbCrLf & oldFolder & vbCrLf & vbCrLf & _
               "Select two different folders.", vbExclamation, "Same Folder"
        Exit Sub
    End If

    GrantMacAccess Array(oldFolder, newFolder)
    Set jobs = BuildBatchJobs(oldFolder, newFolder, wbDashboard.FullName)
    If jobs Is Nothing Then Exit Sub

    RunJobs wbDashboard, jobs, "Batch"
End Sub

' =========================================================================
' [OPEN] 2. USER-FACING MACRO: RUN SINGLE ALERT COMPARISON
' =========================================================================
Public Sub Run_Single_Pair_Comparison()
    Dim oldFile As String, newFile As String
    Dim wbDashboard As Workbook
    Dim oldEcm As String, oldAlert As String, oldKey As String
    Dim newEcm As String, newAlert As String, newKey As String
    Dim jobs As Collection

    Set wbDashboard = ActiveWorkbook
    If wbDashboard Is Nothing Then
        MsgBox "Open the workbook that should hold the dashboard first.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    MsgBox "Please select the OLD / BASELINE transaction file.", vbInformation, "Step 1 of 2: Select Old File"
    oldFile = Pick_Excel_File("Select OLD Transaction File")
    If oldFile = "" Then Exit Sub

    MsgBox "Please select the REVISED / NEW lookback transaction file.", vbInformation, "Step 2 of 2: Select New File"
    newFile = Pick_Excel_File("Select NEW Transaction File")
    If newFile = "" Then Exit Sub

    If StrComp(oldFile, newFile, vbTextCompare) = 0 Then
        MsgBox "The same file was selected as both Old and New.", vbExclamation, "Same File"
        Exit Sub
    End If

    ExtractFileIDs oldFile, oldEcm, oldAlert, oldKey
    ExtractFileIDs newFile, newEcm, newAlert, newKey
    If oldKey <> newKey Then
        If MsgBox("The file names point to different alerts:" & vbCrLf & vbCrLf & _
                  "Old: " & oldEcm & " / " & oldAlert & vbCrLf & _
                  "New: " & newEcm & " / " & newAlert & vbCrLf & vbCrLf & _
                  "Compare them anyway?", vbQuestion + vbYesNo, "Alert Mismatch") <> vbYes Then Exit Sub
    End If

    GrantMacAccess Array(oldFile, newFile)
    Set jobs = New Collection
    jobs.Add Array(newEcm, newAlert, oldFile, newFile, "PAIR")
    RunJobs wbDashboard, jobs, "Single Alert"
End Sub

' =========================================================================
' [RESET] 3. USER-FACING MACRO: RESET / CLEAR DASHBOARD
' =========================================================================
Public Sub Clear_Comparison_Dashboard()
    Dim wb As Workbook
    Dim sheetName As Variant
    Dim ws As Worksheet

    If MsgBox("Clear the Comparison Dashboard and delete all discrepancy sheets?", _
              vbQuestion + vbYesNo, "Confirm Reset") <> vbYes Then Exit Sub

    Set wb = ActiveWorkbook
    Setup_Comparison_Dashboard wb, True

    Application.DisplayAlerts = False
    For Each sheetName In Array(SHEET_ADDED_TXNS, SHEET_DROPPED_TXNS, SHEET_MODIFIED_TXNS)
        Set ws = FindSheet(wb, CStr(sheetName))
        If Not ws Is Nothing Then
            On Error Resume Next
            ws.Delete
            On Error GoTo 0
        End If
    Next sheetName
    Application.DisplayAlerts = True

    FindSheet(wb, SHEET_DASHBOARD).Activate
    MsgBox "Dashboard has been reset.", vbInformation, "Cleared"
End Sub

' =========================================================================
' [EXPORT] 3b. USER-FACING MACRO: EXPORT TO STANDALONE FORMATTED WORKBOOK
' =========================================================================
Public Sub Export_To_New_Workbook()
    Dim wbSource As Workbook, wbNew As Workbook
    Dim wsDash As Worksheet, ws As Worksheet
    Dim defaultFileName As String, savePath As Variant, finalPath As String
    Dim sheetName As Variant
    Dim st As AppState
    Dim errDesc As String

    Set wbSource = ActiveWorkbook
    If wbSource Is Nothing Then Exit Sub
    Set wsDash = FindSheet(wbSource, SHEET_DASHBOARD)
    If wsDash Is Nothing Then
        MsgBox "Comparison Dashboard sheet not found. Please run a comparison first!", vbExclamation, "No Data to Export"
        Exit Sub
    End If

    defaultFileName = "AML_Lookback_Reconciliation_Report_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"

#If Mac Then
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName)
#Else
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, _
                                            FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
                                            Title:="Export AML Comparison Report to New Workbook")
#End If
    If VarType(savePath) = vbBoolean Then Exit Sub      ' user cancelled

    finalPath = Trim$(CStr(savePath))
    If finalPath = "" Then Exit Sub
    If LCase$(Right$(finalPath, 5)) <> ".xlsx" Then finalPath = StripExcelExtension(finalPath) & ".xlsx"

    If StrComp(finalPath, wbSource.FullName, vbTextCompare) = 0 Then
        MsgBox "Choose a different file name - that is the workbook currently open.", vbExclamation, "Export"
        Exit Sub
    End If
    If FileExists(finalPath) Then
        If MsgBox("This file already exists:" & vbCrLf & finalPath & vbCrLf & vbCrLf & "Overwrite it?", _
                  vbQuestion + vbYesNo, "Confirm Overwrite") <> vbYes Then Exit Sub
    End If

    SaveAndSpeedUp st
    On Error GoTo ExportFail

    ' Copy Dashboard to a new workbook, then strip its buttons
    wsDash.Copy
    Set wbNew = ActiveWorkbook
    DeleteButtonShapes wbNew.Worksheets(1)

    ' Copy discrepancy sheets if present
    For Each sheetName In Array(SHEET_ADDED_TXNS, SHEET_DROPPED_TXNS, SHEET_MODIFIED_TXNS)
        Set ws = FindSheet(wbSource, CStr(sheetName))
        If Not ws Is Nothing Then ws.Copy After:=wbNew.Worksheets(wbNew.Worksheets.Count)
    Next sheetName

    wbNew.Worksheets(1).Activate
    wbNew.SaveAs Filename:=finalPath, FileFormat:=51     ' 51 = xlOpenXMLWorkbook (.xlsx)

    RestoreApp st
    MsgBox "AML Lookback Report successfully exported to:" & vbCrLf & vbCrLf & _
           finalPath, vbInformation, "Export Successful"
    Exit Sub

ExportFail:
    errDesc = Err.Description
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    On Error GoTo 0
    RestoreApp st
    MsgBox "Export failed:" & vbCrLf & errDesc, vbCritical, "Export Error"
End Sub

' =========================================================================
' [VIEW] 4. USER-FACING MACROS: VIEW GRANULAR DISCREPANCIES
' =========================================================================
Public Sub View_Added_Transactions()
    ShowSheet SHEET_ADDED_TXNS
End Sub

Public Sub View_Dropped_Transactions()
    ShowSheet SHEET_DROPPED_TXNS
End Sub

Public Sub View_Modified_Transactions()
    ShowSheet SHEET_MODIFIED_TXNS
End Sub

Private Sub ShowSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = FindSheet(ActiveWorkbook, sheetName)
    If ws Is Nothing Then
        MsgBox "The sheet '" & sheetName & "' does not exist yet. Please run a comparison first!", vbExclamation, "No Audit Data"
    Else
        ws.Activate
    End If
End Sub

' =========================================================================
' [ENGINE] 5. JOB BUILDING (BATCH FOLDER PAIRING)
' =========================================================================
' A job is Array(ecmID, alertID, oldPath, newPath, kind)
' kind: PAIR | MISSING_NEW | MISSING_OLD | DUP_OLD | DUP_NEW
Private Function BuildBatchJobs(ByVal oldDir As String, ByVal newDir As String, ByVal excludePath As String) As Collection
    Dim oldFiles As Collection, newFiles As Collection
    Dim oldMap As Object, newMap As Object, seen As Object
    Dim keyOrder As Collection, dupJobs As Collection, jobs As Collection
    Dim fPath As Variant, mKey As Variant
    Dim ecmID As String, alertID As String, matchKey As String
    Dim oldPath As String, newPath As String
    Dim keysArr As Variant, i As Long

    Set oldFiles = ListExcelFiles(oldDir, excludePath)
    Set newFiles = ListExcelFiles(newDir, excludePath)

    If oldFiles.Count = 0 Then
        MsgBox "No Excel files (.xlsx / .xlsm / .xls / .xlsb) found in the Old folder:" & vbCrLf & oldDir, vbExclamation, "No Files Found"
        Exit Function
    End If
    If newFiles.Count = 0 Then
        MsgBox "No Excel files (.xlsx / .xlsm / .xls / .xlsb) found in the New folder:" & vbCrLf & newDir, vbExclamation, "No Files Found"
        Exit Function
    End If

    Set oldMap = CreateLookupDict()
    Set newMap = CreateLookupDict()
    Set seen = CreateLookupDict()
    Set keyOrder = New Collection
    Set dupJobs = New Collection

    For Each fPath In oldFiles
        ExtractFileIDs CStr(fPath), ecmID, alertID, matchKey
        If DictExists(oldMap, matchKey) Then
            dupJobs.Add Array(ecmID, alertID, CStr(fPath), "", "DUP_OLD")
        Else
            DictAdd oldMap, matchKey, CStr(fPath)
            If Not DictExists(seen, matchKey) Then
                DictAdd seen, matchKey, matchKey
                keyOrder.Add matchKey
            End If
        End If
    Next fPath

    For Each fPath In newFiles
        ExtractFileIDs CStr(fPath), ecmID, alertID, matchKey
        If DictExists(newMap, matchKey) Then
            dupJobs.Add Array(ecmID, alertID, "", CStr(fPath), "DUP_NEW")
        Else
            DictAdd newMap, matchKey, CStr(fPath)
            If Not DictExists(seen, matchKey) Then
                DictAdd seen, matchKey, matchKey
                keyOrder.Add matchKey
            End If
        End If
    Next fPath

    ' Sort match keys for a stable, readable dashboard order
    keysArr = CollectionToArray(keyOrder)
    SortStrings keysArr

    Set jobs = New Collection
    For i = LBound(keysArr) To UBound(keysArr)
        mKey = keysArr(i)
        oldPath = ""
        newPath = ""
        If DictExists(oldMap, CStr(mKey)) Then oldPath = DictGet(oldMap, CStr(mKey))
        If DictExists(newMap, CStr(mKey)) Then newPath = DictGet(newMap, CStr(mKey))

        If oldPath <> "" Then
            ExtractFileIDs oldPath, ecmID, alertID, matchKey
        Else
            ExtractFileIDs newPath, ecmID, alertID, matchKey
        End If

        If oldPath <> "" And newPath <> "" Then
            jobs.Add Array(ecmID, alertID, oldPath, newPath, "PAIR")
        ElseIf oldPath <> "" Then
            jobs.Add Array(ecmID, alertID, oldPath, "", "MISSING_NEW")
        Else
            jobs.Add Array(ecmID, alertID, "", newPath, "MISSING_OLD")
        End If
    Next i

    For Each fPath In dupJobs
        jobs.Add fPath
    Next fPath

    Set BuildBatchJobs = jobs
End Function

' =========================================================================
' [ENGINE] 6. CORE COMPARISON RUNNER (BATCH AND SINGLE)
' =========================================================================
Private Sub RunJobs(ByVal wbDash As Workbook, ByVal jobs As Collection, ByVal modeLabel As String)
    Dim wsDash As Worksheet, wsAdded As Worksheet, wsDropped As Worksheet, wsModified As Worksheet
    Dim st As AppState
    Dim tot As BatchTotals
    Dim job As Variant
    Dim curRow As Long, jobNo As Long
    Dim stepName As String, summary As String

    SaveAndSpeedUp st
    On Error GoTo ErrorHandler

    stepName = "Setting Up Worksheets"
    Set wsDash = Setup_Comparison_Dashboard(wbDash)
    Set wsAdded = GetOrCreateWorksheet(wbDash, SHEET_ADDED_TXNS)
    Set wsDropped = GetOrCreateWorksheet(wbDash, SHEET_DROPPED_TXNS)
    Set wsModified = GetOrCreateWorksheet(wbDash, SHEET_MODIFIED_TXNS)

    FormatAuditSheetHeaders wsAdded, "NEWLY ADDED TRANSACTIONS (IN NEW FILE, NOT IN OLD FILE)", AUDIT_ADDED
    FormatAuditSheetHeaders wsDropped, "DROPPED / EXCLUDED TRANSACTIONS (IN OLD FILE, NOT IN NEW FILE)", AUDIT_DROPPED
    FormatAuditSheetHeaders wsModified, "MODIFIED TRANSACTIONS (SAME TRANSACTION ID, DIFFERENT VALUES)", AUDIT_MODIFIED

    stepName = "Clearing Old Dashboard Rows"
    ClearDashboardTable wsDash

    stepName = "Comparing Alert Pairs"
    curRow = FIRST_DATA_ROW
    For Each job In jobs
        jobNo = jobNo + 1
        Application.StatusBar = "AML comparison " & jobNo & " of " & jobs.Count & ": " & CStr(job(1))
        Select Case CStr(job(4))
            Case "PAIR"
                ProcessPair wsDash, wsAdded, wsDropped, wsModified, curRow, CStr(job(0)), CStr(job(1)), _
                            CStr(job(2)), CStr(job(3)), tot
            Case "MISSING_NEW"
                tot.Unmatched = tot.Unmatched + 1
                WriteStatusRow wsDash, curRow, CStr(job(0)), CStr(job(1)), GetFileName(CStr(job(2))), _
                               "MISSING IN NEW FOLDER", "UNMATCHED - no New file for this alert [WARN]"
            Case "MISSING_OLD"
                tot.Unmatched = tot.Unmatched + 1
                WriteStatusRow wsDash, curRow, CStr(job(0)), CStr(job(1)), "MISSING IN OLD FOLDER", _
                               GetFileName(CStr(job(3))), "UNMATCHED - no Old file for this alert [WARN]"
            Case "DUP_OLD"
                tot.DupFiles = tot.DupFiles + 1
                WriteStatusRow wsDash, curRow, CStr(job(0)), CStr(job(1)), GetFileName(CStr(job(2))), "-", _
                               "SKIPPED - another Old file has the same ECM/Alert ID [WARN]"
            Case "DUP_NEW"
                tot.DupFiles = tot.DupFiles + 1
                WriteStatusRow wsDash, curRow, CStr(job(0)), CStr(job(1)), "-", GetFileName(CStr(job(3))), _
                               "SKIPPED - another New file has the same ECM/Alert ID [WARN]"
        End Select
        curRow = curRow + 1
    Next job

    stepName = "Formatting Dashboard Table"
    FormatBatchTable wsDash, FIRST_DATA_ROW, curRow - 1
    FinishAuditSheet wsAdded, 10, "No added transactions."
    FinishAuditSheet wsDropped, 9, "No dropped transactions."
    FinishAuditSheet wsModified, 6, "No modified transactions."

    stepName = "Updating KPI Cards"
    UpdateKPICards wsDash, tot

    wsDash.Activate
    RestoreApp st

    summary = modeLabel & " Comparison Complete!" & vbCrLf & vbCrLf & _
              "- Alert pairs compared: " & tot.Evaluated & vbCrLf & _
              "- Unmatched files: " & tot.Unmatched & "   Duplicate files: " & tot.DupFiles & "   Errors: " & tot.Errors & vbCrLf & _
              "- Alerts with increased count: " & tot.IncreasedCount & vbCrLf & _
              "- Transactions added: " & tot.Added & "   dropped: " & tot.Dropped & "   modified: " & tot.Modified & vbCrLf & _
              "- Newly alerted transactions: " & (tot.AddedAlerted + tot.FlipToYes) & _
              " (added " & tot.AddedAlerted & ", flag No->Yes " & tot.FlipToYes & ")" & vbCrLf & _
              "- New counterparties: " & tot.NewCP & vbCrLf & _
              "- Gross volume delta: " & Format(tot.VolumeDelta, "#,##0.00") & " " & tot.Currencies
    If InStr(tot.Currencies, ",") > 0 Then summary = summary & vbCrLf & vbCrLf & "WARNING: mixed currencies - volume totals are not comparable."
    If tot.Errors > 0 Then summary = summary & vbCrLf & vbCrLf & "Some files could not be compared - see rows marked [ERROR]."
    MsgBox summary, IIf(tot.Errors > 0, vbExclamation, vbInformation), "Reconciliation Finished"
    Exit Sub

ErrorHandler:
    RestoreApp st
    MsgBox "Error during comparison (" & stepName & "):" & vbCrLf & Err.Description, vbCritical, "Error " & Err.Number
End Sub

' =========================================================================
' [SEARCH] 7. RECONCILE ONE OLD/NEW WORKBOOK PAIR
' =========================================================================
Private Sub ProcessPair(ByVal wsDash As Worksheet, ByVal wsAdded As Worksheet, ByVal wsDropped As Worksheet, _
                        ByVal wsModified As Worksheet, ByVal rowIdx As Long, _
                        ByVal ecmID As String, ByVal alertID As String, _
                        ByVal oldFilePath As String, ByVal newFilePath As String, _
                        ByRef tot As BatchTotals)
    Dim oldSet As TxnSet, newSet As TxnSet
    Dim pr As PairResult
    Dim oldCP As Object
    Dim countDelta As Long
    Dim statusStr As String, stage As String
    Dim rowVals(1 To 1, 1 To 17) As Variant

    On Error GoTo PairFail

    ' Each file is read into memory and closed before the next one is opened,
    ' so Old and New files may share the same file name.
    stage = "reading Old file"
    Set oldCP = CreateLookupDict()
    LoadTxnFile oldFilePath, oldSet, oldCP

    stage = "reading New file"
    LoadTxnFile newFilePath, newSet, Nothing

    stage = "comparing transactions"
    CompareSets wsAdded, wsDropped, wsModified, ecmID, alertID, oldSet, newSet, oldCP, pr

    stage = "writing dashboard row"
    countDelta = newSet.UniqueCount - oldSet.UniqueCount
    statusStr = BuildStatus(oldSet, newSet, pr, countDelta)

    rowVals(1, 1) = ecmID
    rowVals(1, 2) = alertID
    rowVals(1, 3) = GetFileName(oldFilePath) & "  [" & oldSet.SheetName & "]"
    rowVals(1, 4) = GetFileName(newFilePath) & "  [" & newSet.SheetName & "]"
    rowVals(1, 5) = FormatPeriod(oldSet.MinDate, oldSet.MaxDate, oldFilePath)
    rowVals(1, 6) = FormatPeriod(newSet.MinDate, newSet.MaxDate, newFilePath)
    rowVals(1, 7) = oldSet.UniqueCount
    rowVals(1, 8) = newSet.UniqueCount
    rowVals(1, 9) = countDelta
    If oldSet.UniqueCount > 0 Then
        rowVals(1, 10) = countDelta / oldSet.UniqueCount
    Else
        rowVals(1, 10) = "n/a"
    End If
    If oldSet.HasAlert And newSet.HasAlert Then
        rowVals(1, 11) = oldSet.AlertCount
        rowVals(1, 12) = newSet.AlertCount
        rowVals(1, 13) = newSet.AlertCount - oldSet.AlertCount
    Else
        rowVals(1, 11) = IIf(oldSet.HasAlert, oldSet.AlertCount, "n/a")
        rowVals(1, 12) = IIf(newSet.HasAlert, newSet.AlertCount, "n/a")
        rowVals(1, 13) = "n/a"
    End If
    If oldSet.HasAmount And newSet.HasAmount Then
        rowVals(1, 14) = oldSet.TotalAmt
        rowVals(1, 15) = newSet.TotalAmt
        rowVals(1, 16) = newSet.TotalAmt - oldSet.TotalAmt
        tot.VolumeDelta = tot.VolumeDelta + (newSet.TotalAmt - oldSet.TotalAmt)
    Else
        rowVals(1, 14) = IIf(oldSet.HasAmount, oldSet.TotalAmt, "n/a")
        rowVals(1, 15) = IIf(newSet.HasAmount, newSet.TotalAmt, "n/a")
        rowVals(1, 16) = "n/a"
    End If
    rowVals(1, 17) = statusStr

    wsDash.Cells(rowIdx, 2).Resize(1, 6).NumberFormat = "@"
    wsDash.Cells(rowIdx, 18).NumberFormat = "@"
    wsDash.Cells(rowIdx, 2).Resize(1, DASH_COLS).Value = rowVals

    ' Totals
    tot.Evaluated = tot.Evaluated + 1
    If countDelta > 0 Then tot.IncreasedCount = tot.IncreasedCount + 1
    If pr.AddedCount > 0 Or pr.DroppedCount > 0 Or pr.ModifiedTxns > 0 Then tot.ChangedPairs = tot.ChangedPairs + 1
    tot.Added = tot.Added + pr.AddedCount
    tot.Dropped = tot.Dropped + pr.DroppedCount
    tot.AddedAlerted = tot.AddedAlerted + pr.AddedAlerted
    tot.FlipToYes = tot.FlipToYes + pr.FlipToYes
    tot.Modified = tot.Modified + pr.ModifiedTxns
    tot.NewCP = tot.NewCP + pr.NewCPDistinct
    tot.Currencies = MergeList(tot.Currencies, oldSet.Currencies)
    tot.Currencies = MergeList(tot.Currencies, newSet.Currencies)
    Exit Sub

PairFail:
    tot.Errors = tot.Errors + 1
    WriteStatusRow wsDash, rowIdx, ecmID, alertID, GetFileName(oldFilePath), GetFileName(newFilePath), _
                   "ERROR while " & stage & ": " & Err.Description & " [ERROR]"
End Sub

Private Function BuildStatus(ByRef oldSet As TxnSet, ByRef newSet As TxnSet, ByRef pr As PairResult, ByVal countDelta As Long) As String
    Dim parts As String, sev As Long, ccyAll As String
    ' sev: 0 = OK, 1 = WARN, 2 = ALERT

    If countDelta > 0 Then
        AddPart parts, "NET COUNT +" & countDelta
        If sev < 1 Then sev = 1
    ElseIf countDelta < 0 Then
        AddPart parts, "NET COUNT " & countDelta
        If sev < 1 Then sev = 1
    End If

    If pr.AddedCount > 0 Then
        If pr.AddedAlerted > 0 Then
            AddPart parts, "ADDED " & pr.AddedCount & " (" & pr.AddedAlerted & " ALERTED)"
            sev = 2
        Else
            AddPart parts, "ADDED " & pr.AddedCount
            If sev < 1 Then sev = 1
        End If
    End If

    If pr.DroppedCount > 0 Then
        If pr.DroppedAlerted > 0 Then
            AddPart parts, "DROPPED " & pr.DroppedCount & " (" & pr.DroppedAlerted & " ALERTED)"
        Else
            AddPart parts, "DROPPED " & pr.DroppedCount
        End If
        If sev < 1 Then sev = 1
    End If

    If pr.FlipToYes > 0 Then
        AddPart parts, "ALERT FLAG NO->YES " & pr.FlipToYes
        sev = 2
    End If
    If pr.FlipToNo > 0 Then
        AddPart parts, "ALERT FLAG YES->NO " & pr.FlipToNo
        If sev < 1 Then sev = 1
    End If
    If pr.FieldChangedTxns > 0 Then
        AddPart parts, "FIELDS CHANGED ON " & pr.FieldChangedTxns & " TXNS"
        If sev < 1 Then sev = 1
    End If

    If pr.NewCPDistinct > 0 Then
        AddPart parts, "NEW CP " & pr.NewCPDistinct & " (" & pr.NewCPTxns & " TXNS)"
        sev = 2
    End If
    If pr.CPUnknown > 0 Then
        AddPart parts, "CP CHECK N/A FOR " & pr.CPUnknown & " TXNS"
        If sev < 1 Then sev = 1
    End If

    If oldSet.DupCount > 0 Or newSet.DupCount > 0 Then
        AddPart parts, "DUPLICATE TXN IDS (old " & oldSet.DupCount & ", new " & newSet.DupCount & ")"
        If sev < 1 Then sev = 1
    End If
    If oldSet.BadDates > 0 Or newSet.BadDates > 0 Then
        AddPart parts, "UNREADABLE DATES (old " & oldSet.BadDates & ", new " & newSet.BadDates & ")"
        If sev < 1 Then sev = 1
    End If
    If oldSet.MissingCols <> "" Then
        AddPart parts, "OLD FILE MISSING: " & oldSet.MissingCols
        If sev < 1 Then sev = 1
    End If
    If newSet.MissingCols <> "" Then
        AddPart parts, "NEW FILE MISSING: " & newSet.MissingCols
        If sev < 1 Then sev = 1
    End If

    ccyAll = MergeList(oldSet.Currencies, newSet.Currencies)
    If InStr(ccyAll, ",") > 0 Then
        AddPart parts, "MIXED CCY (" & ccyAll & ")"
        If sev < 1 Then sev = 1
    End If

    If parts = "" Then parts = "NO CHANGE"
    Select Case sev
        Case 0: BuildStatus = parts & " [OK]"
        Case 1: BuildStatus = parts & " [WARN]"
        Case Else: BuildStatus = parts & " [ALERT]"
    End Select
End Function

Private Sub AddPart(ByRef parts As String, ByVal s As String)
    If parts = "" Then
        parts = s
    Else
        parts = parts & " | " & s
    End If
End Sub

' =========================================================================
' [METRICS] 8. LOAD AND READ A TRANSACTION FILE
' =========================================================================
Private Sub LoadTxnFile(ByVal fPath As String, ByRef ts As TxnSet, ByVal cpDict As Object)
    Dim wb As Workbook, ws As Worksheet
    Dim openedByUs As Boolean
    Dim errNum As Long, errDesc As String

    On Error GoTo LoadFail
    Set wb = GetOrOpenWorkbook(fPath, openedByUs)
    Set ws = FindTransactionSheet(wb)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 520, , "no sheet with a 'Transaction ID' header in " & GetFileName(fPath)
    End If

    ReadTxnSheet ws, ts
    If Not cpDict Is Nothing Then HarvestCounterparties wb, ts, cpDict

    If openedByUs Then wb.Close SaveChanges:=False
    Exit Sub

LoadFail:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If openedByUs Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise errNum, , errDesc
End Sub

Private Sub ReadTxnSheet(ByVal ws As Worksheet, ByRef ts As TxnSet)
    Dim cm As ColMap
    Dim hdrRow As Long, lastRow As Long, lastCol As Long, nRows As Long, i As Long
    Dim hdr As Variant
    Dim colTx As Variant, colAl As Variant, colDt As Variant, colAm As Variant, colAc As Variant
    Dim colDe As Variant, colCp As Variant, colDc As Variant, colCy As Variant, colOr As Variant, colBe As Variant
    Dim txDisp As String, txK As String, alertStr As String
    Dim dt As Variant, amt As Variant, acc As String, descTxt As String
    Dim cp As String, drcr As String, ccy As String, origName As String, benName As String
    Dim d As Date, a As Double, ok As Boolean
    Dim ccyDict As Object

    Set ts.Map = CreateLookupDict()
    Set ccyDict = CreateLookupDict()
    ts.SheetName = ws.Name
    ts.MinDate = Empty
    ts.MaxDate = Empty

    hdrRow = FindHeaderRow(ws, FLD_TXID)
    If hdrRow = 0 Then Err.Raise vbObjectError + 521, , "no 'Transaction ID' header found on sheet '" & ws.Name & "'"
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedCol(ws)

    hdr = ReadBlock(ws, hdrRow, 1, hdrRow, lastCol)
    DetectColumns hdr, lastCol, cm

    ts.HasAlert = (cm.Alert > 0)
    ts.HasDate = (cm.TxDate > 0)
    ts.HasAmount = (cm.Amount > 0)
    ts.HasAccount = (cm.Account > 0)
    ts.HasDrCr = (cm.DrCr > 0)
    ts.HasCcy = (cm.Ccy > 0)
    ts.HasCP = (cm.CP > 0) Or (cm.DrCr > 0 And (cm.OrigName > 0 Or cm.BenName > 0))
    If cm.CP > 0 Then
        ts.CPSource = "Counterparty column"
    ElseIf ts.HasCP Then
        ts.CPSource = "derived from Dr/Cr + Originator/Beneficiary Name"
    Else
        ts.CPSource = "not available"
    End If

    If Not ts.HasAlert Then AddListItem ts.MissingCols, "Is Alerted"
    If Not ts.HasDate Then AddListItem ts.MissingCols, "Transaction Date"
    If Not ts.HasAmount Then AddListItem ts.MissingCols, "Amount"

    If lastRow <= hdrRow Then Exit Sub
    nRows = lastRow - hdrRow

    colTx = ReadBlock(ws, hdrRow + 1, cm.TxID, lastRow, cm.TxID)
    If cm.Alert > 0 Then colAl = ReadBlock(ws, hdrRow + 1, cm.Alert, lastRow, cm.Alert)
    If cm.TxDate > 0 Then colDt = ReadBlock(ws, hdrRow + 1, cm.TxDate, lastRow, cm.TxDate)
    If cm.Amount > 0 Then colAm = ReadBlock(ws, hdrRow + 1, cm.Amount, lastRow, cm.Amount)
    If cm.Account > 0 Then colAc = ReadBlock(ws, hdrRow + 1, cm.Account, lastRow, cm.Account)
    If cm.Desc > 0 Then colDe = ReadBlock(ws, hdrRow + 1, cm.Desc, lastRow, cm.Desc)
    If cm.CP > 0 Then colCp = ReadBlock(ws, hdrRow + 1, cm.CP, lastRow, cm.CP)
    If cm.DrCr > 0 Then colDc = ReadBlock(ws, hdrRow + 1, cm.DrCr, lastRow, cm.DrCr)
    If cm.Ccy > 0 Then colCy = ReadBlock(ws, hdrRow + 1, cm.Ccy, lastRow, cm.Ccy)
    If cm.OrigName > 0 Then colOr = ReadBlock(ws, hdrRow + 1, cm.OrigName, lastRow, cm.OrigName)
    If cm.BenName > 0 Then colBe = ReadBlock(ws, hdrRow + 1, cm.BenName, lastRow, cm.BenName)

    For i = 1 To nRows
        txDisp = CellText(colTx(i, 1))
        If txDisp <> "" Then
            ts.RowCount = ts.RowCount + 1
            txK = TxKey(txDisp)
            If DictExists(ts.Map, txK) Then
                ts.DupCount = ts.DupCount + 1
            Else
                ' Alert flag
                If cm.Alert > 0 Then
                    If IsYes(CellText(colAl(i, 1))) Then
                        alertStr = "Yes"
                        ts.AlertCount = ts.AlertCount + 1
                    Else
                        alertStr = "No"
                    End If
                Else
                    alertStr = "N/A"
                End If

                ' Date
                dt = Empty
                If cm.TxDate > 0 Then
                    d = ParseDateValue(colDt(i, 1), ok)
                    If ok Then
                        dt = d
                        If IsEmpty(ts.MinDate) Then
                            ts.MinDate = d
                            ts.MaxDate = d
                        Else
                            If d < ts.MinDate Then ts.MinDate = d
                            If d > ts.MaxDate Then ts.MaxDate = d
                        End If
                    ElseIf CellText(colDt(i, 1)) <> "" Then
                        dt = CellText(colDt(i, 1))
                        ts.BadDates = ts.BadDates + 1
                    End If
                End If

                ' Amount (gross volume uses the absolute value)
                amt = Empty
                If cm.Amount > 0 Then
                    a = ParseAmount(colAm(i, 1), ok)
                    If ok Then
                        amt = a
                        ts.TotalAmt = ts.TotalAmt + Abs(a)
                    End If
                End If

                acc = ""
                descTxt = ""
                cp = ""
                drcr = ""
                ccy = ""
                origName = ""
                benName = ""
                If cm.Account > 0 Then acc = CellText(colAc(i, 1))
                If cm.Desc > 0 Then descTxt = CellText(colDe(i, 1))
                If cm.CP > 0 Then cp = CellText(colCp(i, 1))
                If cm.DrCr > 0 Then drcr = UCase$(CellText(colDc(i, 1)))
                If cm.Ccy > 0 Then ccy = UCase$(CellText(colCy(i, 1)))
                If cm.OrigName > 0 Then origName = CellText(colOr(i, 1))
                If cm.BenName > 0 Then benName = CellText(colBe(i, 1))
                If cp = "" Then cp = DeriveCounterparty(drcr, origName, benName)

                If ccy <> "" Then
                    If Not DictExists(ccyDict, ccy) Then DictAdd ccyDict, ccy, ccy
                End If

                DictAdd ts.Map, txK, Array(txDisp, alertStr, dt, amt, acc, descTxt, cp, drcr, ccy)
            End If
        End If
    Next i

    ts.UniqueCount = DictCount(ts.Map)
    ts.Currencies = JoinSortedKeys(ccyDict)
End Sub

Private Function DeriveCounterparty(ByVal drcr As String, ByVal origName As String, ByVal benName As String) As String
    ' Credit into the account: counterparty is the originator (sender).
    ' Debit out of the account: counterparty is the beneficiary (receiver).
    Select Case UCase$(Trim$(drcr))
        Case "CR", "C", "CREDIT", "CRDT", "CRED"
            DeriveCounterparty = origName
        Case "DR", "D", "DEBIT", "DBIT", "DEB"
            DeriveCounterparty = benName
        Case Else
            DeriveCounterparty = ""
    End Select
End Function

' =========================================================================
' [AUDIT] 9. COMPARE TWO TRANSACTION SETS AND WRITE AUDIT SHEETS
' =========================================================================
Private Sub CompareSets(ByVal wsAdded As Worksheet, ByVal wsDropped As Worksheet, ByVal wsModified As Worksheet, _
                        ByVal ecmID As String, ByVal alertID As String, _
                        ByRef oldSet As TxnSet, ByRef newSet As TxnSet, ByVal oldCP As Object, _
                        ByRef pr As PairResult)
    Dim txKeys As Variant, i As Long, k As String
    Dim rec As Variant, oRec As Variant, kv As Variant
    Dim addedKeys As Collection, droppedKeys As Collection, modRows As Collection
    Dim newCPSeen As Object
    Dim outArr() As Variant
    Dim n As Long, r As Long, c As Long, startRow As Long
    Dim cpFlag As String, cpKey As String
    Dim isNew As Boolean, changedAny As Boolean, changedOther As Boolean, oldHasCPData As Boolean

    Set addedKeys = New Collection
    Set droppedKeys = New Collection
    Set modRows = New Collection
    Set newCPSeen = CreateLookupDict()
    oldHasCPData = oldSet.HasCP Or (DictCount(oldCP) > 0)

    ' New vs Old: added and modified
    txKeys = DictKeys(newSet.Map)
    For i = LBound(txKeys) To UBound(txKeys)
        k = CStr(txKeys(i))
        If DictExists(oldSet.Map, k) Then
            rec = DictGet(newSet.Map, k)
            oRec = DictGet(oldSet.Map, k)
            DiffRecords ecmID, alertID, oRec, rec, oldSet, newSet, modRows, changedAny, changedOther, pr
            If changedAny Then pr.ModifiedTxns = pr.ModifiedTxns + 1
            If changedOther Then pr.FieldChangedTxns = pr.FieldChangedTxns + 1
        Else
            addedKeys.Add k
        End If
    Next i

    ' Old vs New: dropped
    txKeys = DictKeys(oldSet.Map)
    For i = LBound(txKeys) To UBound(txKeys)
        k = CStr(txKeys(i))
        If Not DictExists(newSet.Map, k) Then droppedKeys.Add k
    Next i

    ' --- Added transactions ---
    n = addedKeys.Count
    pr.AddedCount = n
    If n > 0 Then
        ReDim outArr(1 To n, 1 To 10)
        r = 0
        For Each kv In addedKeys
            r = r + 1
            rec = DictGet(newSet.Map, CStr(kv))
            If rec(F_ALERT) = "Yes" Then pr.AddedAlerted = pr.AddedAlerted + 1

            cpFlag = CPStatus(CStr(rec(F_CP)), oldHasCPData, newSet.HasCP, oldCP, isNew)
            If isNew Then
                pr.NewCPTxns = pr.NewCPTxns + 1
                cpKey = NormCP(CStr(rec(F_CP)))
                If Not DictExists(newCPSeen, cpKey) Then
                    DictAdd newCPSeen, cpKey, cpKey
                    pr.NewCPDistinct = pr.NewCPDistinct + 1
                End If
            ElseIf Left$(cpFlag, 7) = "Unknown" Then
                pr.CPUnknown = pr.CPUnknown + 1
            End If

            outArr(r, 1) = ecmID
            outArr(r, 2) = alertID
            outArr(r, 3) = rec(F_TXID)
            outArr(r, 4) = rec(F_ALERT)
            outArr(r, 5) = rec(F_DATE)
            outArr(r, 6) = rec(F_AMT)
            outArr(r, 7) = rec(F_ACC)
            outArr(r, 8) = rec(F_DESC)
            outArr(r, 9) = rec(F_CP)
            outArr(r, 10) = cpFlag
        Next kv

        startRow = WriteAuditBlock(wsAdded, outArr, n, AUDIT_ADDED)
        For r = 1 To n
            If outArr(r, 4) = "Yes" Then HighlightCell wsAdded.Cells(startRow + r - 1, 5), COLOR_ALERT_RED, COLOR_FILL_RED
            If Left$(CStr(outArr(r, 10)), 3) = "Yes" Then
                wsAdded.Cells(startRow + r - 1, 10).Font.Bold = True
                HighlightCell wsAdded.Cells(startRow + r - 1, 11), COLOR_ALERT_RED, COLOR_FILL_RED
            ElseIf Left$(CStr(outArr(r, 10)), 7) = "Unknown" Or Left$(CStr(outArr(r, 10)), 3) = "No " Then
                wsAdded.Cells(startRow + r - 1, 11).Font.Color = COLOR_AMBER
            Else
                wsAdded.Cells(startRow + r - 1, 11).Font.Color = COLOR_SUCCESS_GREEN
            End If
        Next r
    End If

    ' --- Dropped transactions ---
    n = droppedKeys.Count
    pr.DroppedCount = n
    If n > 0 Then
        ReDim outArr(1 To n, 1 To 9)
        r = 0
        For Each kv In droppedKeys
            r = r + 1
            rec = DictGet(oldSet.Map, CStr(kv))
            If rec(F_ALERT) = "Yes" Then pr.DroppedAlerted = pr.DroppedAlerted + 1
            outArr(r, 1) = ecmID
            outArr(r, 2) = alertID
            outArr(r, 3) = rec(F_TXID)
            outArr(r, 4) = rec(F_ALERT)
            outArr(r, 5) = rec(F_DATE)
            outArr(r, 6) = rec(F_AMT)
            outArr(r, 7) = rec(F_ACC)
            outArr(r, 8) = rec(F_DESC)
            outArr(r, 9) = rec(F_CP)
        Next kv

        startRow = WriteAuditBlock(wsDropped, outArr, n, AUDIT_DROPPED)
        For r = 1 To n
            If outArr(r, 4) = "Yes" Then HighlightCell wsDropped.Cells(startRow + r - 1, 5), COLOR_ALERT_RED, COLOR_FILL_RED
        Next r
    End If

    ' --- Modified transactions ---
    n = modRows.Count
    If n > 0 Then
        ReDim outArr(1 To n, 1 To 6)
        r = 0
        For Each kv In modRows
            r = r + 1
            For c = 1 To 6
                outArr(r, c) = kv(c - 1)
            Next c
        Next kv

        startRow = WriteAuditBlock(wsModified, outArr, n, AUDIT_MODIFIED)
        For r = 1 To n
            If outArr(r, 4) = "Is Alerted" And outArr(r, 6) = "Yes" Then
                HighlightCell wsModified.Cells(startRow + r - 1, 7), COLOR_ALERT_RED, COLOR_FILL_RED
            Else
                wsModified.Cells(startRow + r - 1, 7).Font.Color = COLOR_AMBER
            End If
        Next r
    End If
End Sub

Private Sub DiffRecords(ByVal ecmID As String, ByVal alertID As String, ByVal o As Variant, ByVal nw As Variant, _
                        ByRef oldSet As TxnSet, ByRef newSet As TxnSet, ByVal modRows As Collection, _
                        ByRef changedAny As Boolean, ByRef changedOther As Boolean, ByRef pr As PairResult)
    Dim tx As String, oldCPKey As String, newCPKey As String

    changedAny = False
    changedOther = False
    tx = CStr(nw(F_TXID))

    If oldSet.HasAlert And newSet.HasAlert Then
        If o(F_ALERT) <> nw(F_ALERT) Then
            modRows.Add Array(ecmID, alertID, tx, "Is Alerted", o(F_ALERT), nw(F_ALERT))
            If nw(F_ALERT) = "Yes" Then pr.FlipToYes = pr.FlipToYes + 1 Else pr.FlipToNo = pr.FlipToNo + 1
            changedAny = True
        End If
    End If

    If oldSet.HasAmount And newSet.HasAmount Then
        If AmountsDiffer(o(F_AMT), nw(F_AMT)) Then
            modRows.Add Array(ecmID, alertID, tx, "Amount", ValText(o(F_AMT)), ValText(nw(F_AMT)))
            changedAny = True
            changedOther = True
        End If
    End If

    If oldSet.HasDate And newSet.HasDate Then
        If ValText(o(F_DATE)) <> ValText(nw(F_DATE)) Then
            modRows.Add Array(ecmID, alertID, tx, "Transaction Date", ValText(o(F_DATE)), ValText(nw(F_DATE)))
            changedAny = True
            changedOther = True
        End If
    End If

    If oldSet.HasDrCr And newSet.HasDrCr Then
        If o(F_DRCR) <> nw(F_DRCR) Then
            modRows.Add Array(ecmID, alertID, tx, "Dr Cr", ValText(o(F_DRCR)), ValText(nw(F_DRCR)))
            changedAny = True
            changedOther = True
        End If
    End If

    If oldSet.HasCcy And newSet.HasCcy Then
        If o(F_CCY) <> nw(F_CCY) Then
            modRows.Add Array(ecmID, alertID, tx, "Currency", ValText(o(F_CCY)), ValText(nw(F_CCY)))
            changedAny = True
            changedOther = True
        End If
    End If

    If oldSet.HasAccount And newSet.HasAccount Then
        If StrComp(CStr(o(F_ACC)), CStr(nw(F_ACC)), vbTextCompare) <> 0 Then
            modRows.Add Array(ecmID, alertID, tx, "Account No", ValText(o(F_ACC)), ValText(nw(F_ACC)))
            changedAny = True
            changedOther = True
        End If
    End If

    If oldSet.HasCP And newSet.HasCP Then
        oldCPKey = NormCP(CStr(o(F_CP)))
        newCPKey = NormCP(CStr(nw(F_CP)))
        If oldCPKey <> "" And newCPKey <> "" And oldCPKey <> newCPKey Then
            modRows.Add Array(ecmID, alertID, tx, "Counterparty", ValText(o(F_CP)), ValText(nw(F_CP)))
            changedAny = True
            changedOther = True
        End If
    End If
End Sub

Private Function AmountsDiffer(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmpty(a) And IsEmpty(b) Then
        AmountsDiffer = False
    ElseIf IsEmpty(a) Or IsEmpty(b) Then
        AmountsDiffer = True
    Else
        AmountsDiffer = (Abs(CDbl(a) - CDbl(b)) > AMOUNT_TOLERANCE)
    End If
End Function

Private Function CPStatus(ByVal cp As String, ByVal oldHasCPData As Boolean, ByVal newHasCP As Boolean, _
                          ByVal oldCP As Object, ByRef isNew As Boolean) As String
    Dim cpKey As String, oldName As String
    isNew = False

    If Not newHasCP Then
        CPStatus = "Unknown (no CP data in New file)"
        Exit Function
    End If
    If Not oldHasCPData Then
        CPStatus = "Unknown (no CP data in Old file)"
        Exit Function
    End If

    cpKey = NormCP(cp)
    If cpKey = "" Then
        CPStatus = "Unknown (blank CP)"
    ElseIf DictExists(oldCP, cpKey) Then
        oldName = CStr(DictGet(oldCP, cpKey))
        If StrComp(Trim$(oldName), Trim$(cp), vbTextCompare) = 0 Then
            CPStatus = "No"
        Else
            CPStatus = "No (in Old file as: " & oldName & ")"
        End If
    Else
        isNew = True
        CPStatus = "Yes [NEW CP]"
    End If
End Function

' =========================================================================
' [CP] COLLECT KNOWN COUNTERPARTIES FROM THE OLD FILE
' =========================================================================
' Sources: the primary transaction sheet (column or derived), any other sheet
' with a Counterparty column (e.g. "CP Selection", "DeDupe"), and any pivot
' table field named like "Counterparty". Keys are normalised names.
Private Sub HarvestCounterparties(ByVal wb As Workbook, ByRef ts As TxnSet, ByVal cpDict As Object)
    Dim txKeys As Variant, i As Long, rec As Variant
    Dim ws As Worksheet
    Dim hdrRow As Long, lastRow As Long, lastCol As Long, c As Long
    Dim hdr As Variant, colVals As Variant

    txKeys = DictKeys(ts.Map)
    For i = LBound(txKeys) To UBound(txKeys)
        rec = DictGet(ts.Map, CStr(txKeys(i)))
        AddCP cpDict, CStr(rec(F_CP))
    Next i

    For Each ws In wb.Worksheets
        If ws.Name <> ts.SheetName Then
            hdrRow = FindHeaderRow(ws, FLD_CP)
            If hdrRow > 0 Then
                lastRow = LastUsedRow(ws)
                lastCol = LastUsedCol(ws)
                If lastRow > hdrRow Then
                    hdr = ReadBlock(ws, hdrRow, 1, hdrRow, lastCol)
                    c = GetFieldCol(hdr, 1, lastCol, FLD_CP)
                    If c > 0 Then
                        colVals = ReadBlock(ws, hdrRow + 1, c, lastRow, c)
                        For i = 1 To UBound(colVals, 1)
                            AddCP cpDict, CellText(colVals(i, 1))
                        Next i
                    End If
                End If
            End If
        End If
    Next ws

    HarvestPivotCounterparties wb, cpDict
End Sub

Private Sub HarvestPivotCounterparties(ByVal wb As Workbook, ByVal cpDict As Object)
    Dim ws As Worksheet, pt As Object, pf As Object
    On Error GoTo Done
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            For Each pf In pt.PivotFields
                If IsCounterpartyPivotField(pf) Then AddPivotItems pf, cpDict
            Next pf
        Next pt
    Next ws
Done:
End Sub

Private Function IsCounterpartyPivotField(ByVal pf As Object) As Boolean
    On Error GoTo Done
    IsCounterpartyPivotField = (HeaderScore(NormText(CStr(pf.SourceName)), FLD_CP) > 0)
Done:
End Function

Private Sub AddPivotItems(ByVal pf As Object, ByVal cpDict As Object)
    Dim pi As Object
    On Error GoTo Done
    For Each pi In pf.PivotItems
        AddCP cpDict, CStr(pi.Name)
    Next pi
Done:
End Sub

Private Sub AddCP(ByVal cpDict As Object, ByVal cpName As String)
    Dim cpKey As String
    cpKey = NormCP(cpName)
    If cpKey = "" Or cpKey = "BLANK" Then Exit Sub
    If Not DictExists(cpDict, cpKey) Then DictAdd cpDict, cpKey, Trim$(cpName)
End Sub

' Normalises a counterparty name: upper case, SWIFT line tags ("1/NAME")
' removed, "&" = "AND", punctuation removed, spaces collapsed.
Private Function NormCP(ByVal s As String) As String
    s = UCase$(Trim$(Replace(s, Chr(160), " ")))
    Do While Len(s) >= 2
        If Not (Left$(s, 1) Like "#" And Mid$(s, 2, 1) = "/") Then Exit Do
        s = Trim$(Mid$(s, 3))
    Loop
    s = Replace(s, "&", " AND ")
    NormCP = UCase$(NormText(s))
End Function

' =========================================================================
' [UI] 10. DASHBOARD FORMATTING & BUTTON SETUP
' =========================================================================
Public Function Setup_Comparison_Dashboard(ByVal wb As Workbook, Optional ByVal forceReset As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Dim headers As Variant
    Dim c As Long

    Set ws = GetOrCreateWorksheet(wb, SHEET_DASHBOARD, True)

    ' Re-use the existing dashboard only if it was built by this version
    If Not forceReset And CellText(ws.Range("A1").Value2) = DASH_VERSION Then
        Set Setup_Comparison_Dashboard = ws
        Exit Function
    End If

    ws.Cells.UnMerge
    ws.Cells.Clear
    DeleteButtonShapes ws, True

    ' Version marker (hidden by number format)
    ws.Range("A1").Value = DASH_VERSION
    ws.Range("A1").NumberFormat = ";;;"

    ' Column Widths (Columns B to R: 17 Data Columns)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 14 ' ECM ID
    ws.Columns("C").ColumnWidth = 18 ' Alert ID
    ws.Columns("D").ColumnWidth = 36 ' Old File [sheet]
    ws.Columns("E").ColumnWidth = 36 ' New File [sheet]
    ws.Columns("F").ColumnWidth = 56 ' Old Period
    ws.Columns("G").ColumnWidth = 56 ' New Period
    ws.Columns("H").ColumnWidth = 11 ' Old Count
    ws.Columns("I").ColumnWidth = 11 ' New Count
    ws.Columns("J").ColumnWidth = 11 ' Delta Count
    ws.Columns("K").ColumnWidth = 10 ' % Change
    ws.Columns("L").ColumnWidth = 11 ' Old Alerted
    ws.Columns("M").ColumnWidth = 11 ' New Alerted
    ws.Columns("N").ColumnWidth = 11 ' Alerted Delta
    ws.Columns("O").ColumnWidth = 16 ' Old Volume
    ws.Columns("P").ColumnWidth = 16 ' New Volume
    ws.Columns("Q").ColumnWidth = 16 ' Volume Delta
    ws.Columns("R").ColumnWidth = 70 ' Status Badge

    ' Row Heights
    ws.Rows("1:3").RowHeight = 12
    ws.Rows(4).RowHeight = 32
    ws.Rows(5).RowHeight = 20
    ws.Rows(6).RowHeight = 10
    ws.Rows("7:9").RowHeight = 22  ' KPI Cards
    ws.Rows(10).RowHeight = 8
    ws.Rows(11).RowHeight = 12
    ws.Rows(12).RowHeight = 30     ' Button Bar
    ws.Rows(13).RowHeight = 12
    ws.Rows(14).RowHeight = 30     ' Table Header

    ' Master Banner Header
    ws.Range("B4:R4").Merge
    With ws.Range("B4")
        .Value = DASH_TITLE
        .Font.Name = FONT_UI
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("B5:R5").Merge
    With ws.Range("B5")
        .Value = "Automated discrepancy verification when review period dates change | Added, dropped and modified transactions, " & _
                 "newly alerted transactions, new counterparties and gross volume shifts"
        .Font.Name = FONT_UI
        .Font.Size = 9.5
        .Font.Color = RGB(100, 116, 139) ' Slate 500
        .Interior.Color = RGB(241, 245, 249) ' Slate 100
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' KPI Tiles (Columns B to R)
    CreateKPITile ws, "B7:D9", "ALERTS EVALUATED", "Alert pairs compared"
    CreateKPITile ws, "E7:G9", "ALERTS WITH INCREASED COUNT", "Lookback count increased"
    CreateKPITile ws, "H7:J9", "TXNS ADDED", "Present in New file only"
    CreateKPITile ws, "K7:M9", "NEWLY ALERTED TXNS", "Added alerted + alert flag No->Yes"
    CreateKPITile ws, "N7:R9", "GROSS VOLUME DELTA", "Sum of absolute amounts, New minus Old"
    ws.Range("N8").NumberFormat = "#,##0.00;-#,##0.00;0.00"

    ' Action Buttons (Row 12)
    CreateActionButton ws, 12, "B", "C", "[RUN] Run Batch Comparison", "Run_Batch_Comparison", COLOR_ACCENT_BLUE
    CreateActionButton ws, 12, "D", "D", "[OPEN] Compare Single Pair", "Run_Single_Pair_Comparison", RGB(79, 70, 229)
    CreateActionButton ws, 12, "E", "E", "[EXPORT] Export to Workbook", "Export_To_New_Workbook", RGB(16, 149, 193)
    CreateActionButton ws, 12, "F", "F", "[VIEW] Added Txns", "View_Added_Transactions", RGB(13, 148, 136)
    CreateActionButton ws, 12, "G", "G", "[VIEW] Dropped Txns", "View_Dropped_Transactions", RGB(13, 148, 136)
    CreateActionButton ws, 12, "H", "K", "[VIEW] Modified Txns", "View_Modified_Transactions", RGB(13, 148, 136)
    CreateActionButton ws, 12, "L", "O", "[RESET] Reset Dashboard", "Clear_Comparison_Dashboard", RGB(100, 116, 139)

    ' Table Header Row
    headers = Array("ECM ID", "Alert ID", "Old File Name [sheet]", "New File Name [sheet]", _
                    "Old Txn Date Range (data | file name)", "New Txn Date Range (data | file name)", _
                    "Old Count", "New Count", "Count Delta", "% Change", "Old Alerted", "New Alerted", "Alerted Delta", _
                    "Old Volume", "New Volume", "Volume Delta", "Reconciliation Status")

    For c = 0 To UBound(headers)
        With ws.Cells(14, c + 2)
            .Value = headers(c)
            .Font.Name = FONT_UI
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

Private Sub CreateKPITile(ByVal ws As Worksheet, ByVal cellRange As String, ByVal title As String, ByVal subTitle As String)
    Dim rng As Range
    Set rng = ws.Range(cellRange)

    rng.Interior.Color = COLOR_CARD_BG
    rng.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=COLOR_BORDER
    rng.Rows(1).Merge
    rng.Rows(2).Merge
    rng.Rows(3).Merge

    ' Title Row
    With rng.Cells(1, 1)
        .Value = title
        .Font.Name = FONT_UI
        .Font.Size = 8.5
        .Font.Bold = True
        .Font.Color = RGB(100, 116, 139)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Value Row
    With rng.Cells(2, 1)
        .Value = 0
        .Font.Name = FONT_UI
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = COLOR_TEXT_DARK
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Subtitle Row
    With rng.Cells(3, 1)
        .Value = subTitle
        .Font.Name = FONT_UI
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(148, 163, 184)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub UpdateKPICards(ByVal ws As Worksheet, ByRef tot As BatchTotals)
    Dim newlyAlerted As Long

    SetKPI ws, "B8", tot.Evaluated, COLOR_TEXT_DARK
    ws.Range("B9").Value = "Compared: " & tot.Evaluated & " | Unmatched: " & tot.Unmatched & _
                           " | Dup files: " & tot.DupFiles & " | Errors: " & tot.Errors
    If tot.Errors > 0 Then ws.Range("B8").Font.Color = COLOR_ALERT_RED

    SetKPI ws, "E8", tot.IncreasedCount, IIf(tot.IncreasedCount > 0, COLOR_AMBER, COLOR_TEXT_DARK)
    ws.Range("E9").Value = "Alerts with any transaction change: " & tot.ChangedPairs

    SetKPI ws, "H8", tot.Added, IIf(tot.Added > 0, COLOR_ACCENT_BLUE, COLOR_TEXT_DARK)
    ws.Range("H9").Value = "Dropped: " & tot.Dropped & " | Modified: " & tot.Modified & _
                           " | Net: " & Format(tot.Added - tot.Dropped, "+0;-0;0")

    newlyAlerted = tot.AddedAlerted + tot.FlipToYes
    SetKPI ws, "K8", newlyAlerted, IIf(newlyAlerted > 0, COLOR_ALERT_RED, COLOR_TEXT_DARK)
    ws.Range("K9").Value = "Added alerted: " & tot.AddedAlerted & " | Flag No->Yes: " & tot.FlipToYes & _
                           " | New CPs: " & tot.NewCP

    If tot.VolumeDelta > 0.005 Then
        SetKPI ws, "N8", tot.VolumeDelta, COLOR_SUCCESS_GREEN
    ElseIf tot.VolumeDelta < -0.005 Then
        SetKPI ws, "N8", tot.VolumeDelta, COLOR_ALERT_RED
    Else
        SetKPI ws, "N8", tot.VolumeDelta, COLOR_TEXT_DARK
    End If
    ws.Range("N8").NumberFormat = "#,##0.00;-#,##0.00;0.00"
    If InStr(tot.Currencies, ",") > 0 Then
        ws.Range("N9").Value = "MIXED CURRENCIES (" & tot.Currencies & ") - totals not comparable"
        ws.Range("N9").Font.Color = COLOR_ALERT_RED
    Else
        ws.Range("N9").Value = "Sum of absolute amounts, New minus Old" & IIf(tot.Currencies <> "", " (" & tot.Currencies & ")", "")
        ws.Range("N9").Font.Color = RGB(148, 163, 184)
    End If
End Sub

Private Sub SetKPI(ByVal ws As Worksheet, ByVal addr As String, ByVal v As Variant, ByVal fontColor As Long)
    ws.Range(addr).Value = v
    ws.Range(addr).Font.Color = fontColor
End Sub

Private Sub CreateActionButton(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal startCol As String, ByVal endCol As String, _
                               ByVal btnText As String, ByVal macroName As String, ByVal bgColor As Long)
    Dim btnRange As Range
    Dim shp As Shape

    Set btnRange = ws.Range(startCol & rowIdx & ":" & endCol & rowIdx)
    Set shp = ws.Shapes.AddShape(5, btnRange.Left + 2, btnRange.Top + 2, btnRange.Width - 4, btnRange.Height - 4) ' 5 = msoShapeRoundedRectangle

    With shp
        .Name = "Btn_" & macroName
        ' Fully qualified so the button still works when this code lives in another workbook
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
        .Adjustments.Item(1) = 0.2
        .Fill.Solid
        .Fill.ForeColor.RGB = bgColor
        .Line.Visible = msoFalse
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = btnText
            .TextRange.Font.Name = FONT_UI
            .TextRange.Font.Size = 10
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Private Sub DeleteButtonShapes(ByVal ws As Worksheet, Optional ByVal allShapes As Boolean = False)
    Dim i As Long
    ' Iterate backwards: deleting inside For Each can skip shapes
    For i = ws.Shapes.Count To 1 Step -1
        If allShapes Or Left$(ws.Shapes(i).Name, 4) = "Btn_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub ClearDashboardTable(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    If lastRow >= FIRST_DATA_ROW Then
        With ws.Range(ws.Cells(FIRST_DATA_ROW, 2), ws.Cells(lastRow, 1 + DASH_COLS))
            .UnMerge
            .Clear
        End With
        ws.Range(ws.Rows(FIRST_DATA_ROW), ws.Rows(lastRow)).RowHeight = ws.StandardHeight
    End If
End Sub

Private Function DashboardHasResults(ByVal ws As Worksheet) As Boolean
    DashboardHasResults = (LastUsedRow(ws) >= FIRST_DATA_ROW)
End Function

Private Sub FormatBatchTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim rngTable As Range
    Dim r As Long
    Dim statusTxt As String

    If endRow < startRow Then Exit Sub
    Set rngTable = ws.Range(ws.Cells(startRow, 2), ws.Cells(endRow, 1 + DASH_COLS))

    With rngTable
        .Font.Name = FONT_UI
        .Font.Size = 9
        .Font.Color = COLOR_TEXT_DARK
        .VerticalAlignment = xlCenter
        .RowHeight = 22
        With .Borders
            .LineStyle = xlContinuous
            .Color = COLOR_BORDER
            .Weight = xlThin
        End With
    End With

    ' Column-level alignment and number formats (applied once to the whole block)
    ws.Range(ws.Cells(startRow, 2), ws.Cells(endRow, 2)).HorizontalAlignment = xlCenter   ' ECM ID
    ws.Range(ws.Cells(startRow, 3), ws.Cells(endRow, 5)).HorizontalAlignment = xlLeft     ' Alert ID, files
    ws.Range(ws.Cells(startRow, 6), ws.Cells(endRow, 7)).HorizontalAlignment = xlCenter   ' Periods
    With ws.Range(ws.Cells(startRow, 8), ws.Cells(endRow, 10))                            ' Counts
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0"
    End With
    With ws.Range(ws.Cells(startRow, 11), ws.Cells(endRow, 11))                           ' % Change
        .HorizontalAlignment = xlRight
        .NumberFormat = "0.00%"
    End With
    With ws.Range(ws.Cells(startRow, 12), ws.Cells(endRow, 14))                           ' Alerted
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0"
    End With
    With ws.Range(ws.Cells(startRow, 15), ws.Cells(endRow, 17))                           ' Volumes
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0.00;-#,##0.00;0.00"
    End With
    With ws.Range(ws.Cells(startRow, 18), ws.Cells(endRow, 18))                           ' Status
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
    End With

    For r = startRow To endRow
        ' Count delta
        If IsNumeric(ws.Cells(r, 10).Value) And Not IsEmpty(ws.Cells(r, 10).Value) Then
            If ws.Cells(r, 10).Value > 0 Then
                HighlightCell ws.Cells(r, 10), COLOR_ALERT_RED, COLOR_FILL_RED
            ElseIf ws.Cells(r, 10).Value < 0 Then
                ws.Cells(r, 10).Font.Bold = True
                ws.Cells(r, 10).Font.Color = COLOR_AMBER
            End If
        End If

        ' Alerted delta
        If IsNumeric(ws.Cells(r, 14).Value) And Not IsEmpty(ws.Cells(r, 14).Value) Then
            If ws.Cells(r, 14).Value > 0 Then
                ws.Cells(r, 14).Font.Bold = True
                ws.Cells(r, 14).Font.Color = COLOR_ALERT_RED
            End If
        End If

        ' Volume delta
        If IsNumeric(ws.Cells(r, 17).Value) And Not IsEmpty(ws.Cells(r, 17).Value) Then
            If ws.Cells(r, 17).Value < 0 Then ws.Cells(r, 17).Font.Color = COLOR_ALERT_RED
        End If

        ' Status badge
        statusTxt = CStr(ws.Cells(r, 18).Value)
        With ws.Cells(r, 18)
            If InStr(statusTxt, "[ERROR]") > 0 Or InStr(statusTxt, "[ALERT]") > 0 Then
                .Font.Color = COLOR_ALERT_RED
                .Interior.Color = COLOR_FILL_RED
            ElseIf InStr(statusTxt, "[WARN]") > 0 Then
                .Font.Color = COLOR_AMBER
                .Interior.Color = COLOR_FILL_AMBER
            ElseIf InStr(statusTxt, "[OK]") > 0 Then
                .Font.Color = COLOR_SUCCESS_GREEN
                .Interior.Color = COLOR_FILL_GREEN
            Else
                .Font.Color = COLOR_TEXT_MUTED
            End If
        End With
    Next r
End Sub

Private Sub HighlightCell(ByVal c As Range, ByVal fontColor As Long, ByVal fillColor As Long)
    c.Font.Bold = True
    c.Font.Color = fontColor
    c.Interior.Color = fillColor
End Sub

Private Sub WriteStatusRow(ByVal wsDash As Worksheet, ByVal rowIdx As Long, ByVal ecmID As String, ByVal alertID As String, _
                           ByVal oldF As String, ByVal newF As String, ByVal statusStr As String)
    Dim rowVals(1 To 1, 1 To 17) As Variant
    Dim c As Long

    rowVals(1, 1) = ecmID
    rowVals(1, 2) = alertID
    rowVals(1, 3) = oldF
    rowVals(1, 4) = newF
    rowVals(1, 5) = "N/A"
    rowVals(1, 6) = "N/A"
    For c = 7 To 16
        rowVals(1, c) = "-"
    Next c
    rowVals(1, 17) = statusStr

    wsDash.Cells(rowIdx, 2).Resize(1, 6).NumberFormat = "@"
    wsDash.Cells(rowIdx, 18).NumberFormat = "@"
    wsDash.Cells(rowIdx, 2).Resize(1, DASH_COLS).Value = rowVals
End Sub

Private Sub FormatAuditSheetHeaders(ByVal ws As Worksheet, ByVal bannerTitle As String, ByVal kind As Long)
    Dim lastColLetter As String
    Dim headers As Variant
    Dim c As Long

    ws.AutoFilterMode = False
    ws.Cells.UnMerge
    ws.Cells.Clear

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 14 ' ECM ID
    ws.Columns("C").ColumnWidth = 18 ' Alert ID
    ws.Columns("D").ColumnWidth = 16 ' Txn ID

    Select Case kind
        Case AUDIT_MODIFIED
            ws.Columns("E").ColumnWidth = 20 ' Field
            ws.Columns("F").ColumnWidth = 40 ' Old Value
            ws.Columns("G").ColumnWidth = 40 ' New Value
            lastColLetter = "G"
            headers = Array("ECM ID", "Alert ID", "Transaction ID", "Field Changed", "Old Value", "New Value")
        Case Else
            ws.Columns("E").ColumnWidth = 12 ' Alerted?
            ws.Columns("F").ColumnWidth = 14 ' Date
            ws.Columns("G").ColumnWidth = 16 ' Amount
            ws.Columns("H").ColumnWidth = 22 ' Account No
            ws.Columns("I").ColumnWidth = 28 ' Description
            ws.Columns("J").ColumnWidth = 34 ' Counterparty
            If kind = AUDIT_ADDED Then
                ws.Columns("K").ColumnWidth = 40 ' New Counterparty?
                lastColLetter = "K"
                headers = Array("ECM ID", "Alert ID", "Transaction ID", "Is Alerted?", "Transaction Date", "Amount", _
                                "Account No", "Transaction Description", "Counterparty Name", "New Counterparty?")
            Else
                lastColLetter = "J"
                headers = Array("ECM ID", "Alert ID", "Transaction ID", "Is Alerted?", "Transaction Date", "Amount", _
                                "Account No", "Transaction Description", "Counterparty Name")
            End If
    End Select

    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 8
    ws.Rows(3).RowHeight = 24

    ws.Range("B1:" & lastColLetter & "1").Merge
    With ws.Range("B1")
        .Value = bannerTitle
        .Font.Name = FONT_UI
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    For c = 0 To UBound(headers)
        With ws.Cells(3, c + 2)
            .Value = headers(c)
            .Font.Name = FONT_UI
            .Font.Size = 9.5
            .Font.Bold = True
            .Font.Color = COLOR_HEADER_TXT
            .Interior.Color = RGB(30, 41, 59) ' Slate 800
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
End Sub

' Writes nRows x nCols of outArr starting at the next free row (column B).
' Returns the first row written.
Private Function WriteAuditBlock(ByVal ws As Worksheet, ByRef outArr() As Variant, ByVal nRows As Long, ByVal kind As Long) As Long
    Dim startRow As Long, nCols As Long
    Dim rng As Range

    startRow = LastUsedRow(ws) + 1
    If startRow < 4 Then startRow = 4
    nCols = UBound(outArr, 2)
    Set rng = ws.Cells(startRow, 2).Resize(nRows, nCols)

    ' Text columns first so IDs / account numbers keep leading zeros and are not turned into numbers
    If kind = AUDIT_MODIFIED Then
        rng.NumberFormat = "@"
    Else
        ws.Cells(startRow, 2).Resize(nRows, 4).NumberFormat = "@"     ' ECM, Alert, Txn ID, Alerted?
        ws.Cells(startRow, 8).Resize(nRows, nCols - 6).NumberFormat = "@" ' Account .. end
    End If

    rng.Value = outArr

    With rng
        .Font.Name = FONT_UI
        .Font.Size = 9
        .Font.Color = COLOR_TEXT_DARK
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        With .Borders
            .LineStyle = xlContinuous
            .Color = COLOR_BORDER
            .Weight = xlThin
        End With
    End With
    ws.Cells(startRow, 2).Resize(nRows, 1).HorizontalAlignment = xlCenter   ' ECM ID
    ws.Cells(startRow, 4).Resize(nRows, 1).HorizontalAlignment = xlCenter   ' Txn ID

    If kind <> AUDIT_MODIFIED Then
        ws.Cells(startRow, 5).Resize(nRows, 2).HorizontalAlignment = xlCenter  ' Alerted?, Date
        ws.Cells(startRow, 6).Resize(nRows, 1).NumberFormat = "yyyy-mm-dd"
        With ws.Cells(startRow, 7).Resize(nRows, 1)
            .NumberFormat = "#,##0.00;-#,##0.00;0.00"
            .HorizontalAlignment = xlRight
        End With
        If kind = AUDIT_ADDED Then ws.Cells(startRow, 11).Resize(nRows, 1).HorizontalAlignment = xlCenter
    End If

    WriteAuditBlock = startRow
End Function

Private Sub FinishAuditSheet(ByVal ws As Worksheet, ByVal nCols As Long, ByVal emptyMsg As String)
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    If lastRow < 4 Then
        With ws.Range("B4")
            .Value = emptyMsg
            .Font.Name = FONT_UI
            .Font.Size = 9
            .Font.Italic = True
            .Font.Color = RGB(148, 163, 184)
        End With
    Else
        ws.Range(ws.Cells(3, 2), ws.Cells(lastRow, 1 + nCols)).AutoFilter
    End If
End Sub

' =========================================================================
' [HELPER] 11. FILES, FOLDERS AND WORKBOOKS (Mac & Windows)
' =========================================================================
Private Function GetOrOpenWorkbook(ByVal fPath As String, ByRef openedByUs As Boolean) As Workbook
    Dim wb As Workbook, fName As String
    Dim errNum As Long, errDesc As String

    openedByUs = False
    fName = GetFileName(fPath)

    ' Already open from this exact path?
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fPath, vbTextCompare) = 0 Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    ' Excel cannot open two workbooks with the same name at once
    Set wb = Nothing
    On Error Resume Next
    Set wb = Application.Workbooks(fName)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Err.Raise vbObjectError + 513, , "a different workbook named '" & fName & "' is already open (" & wb.FullName & "). Close it and run again"
    End If

    ' A dummy password makes protected files fail fast instead of prompting
    On Error Resume Next
    Set wb = Application.Workbooks.Open(Filename:=fPath, UpdateLinks:=0, ReadOnly:=True, _
                                        Password:="~no~password~", IgnoreReadOnlyRecommended:=True, _
                                        AddToMru:=False, Notify:=False)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0

    ' Retry with the basic arguments unless the failure was the password itself
    If wb Is Nothing And InStr(1, errDesc, "password", vbTextCompare) = 0 Then
        On Error Resume Next
        Set wb = Application.Workbooks.Open(Filename:=fPath, UpdateLinks:=0, ReadOnly:=True)
        If Err.Number <> 0 Then
            errNum = Err.Number
            errDesc = Err.Description
        End If
        On Error GoTo 0
    End If
    If wb Is Nothing Then
        Err.Raise vbObjectError + 514, , "could not open '" & fName & "' (" & errNum & ": " & errDesc & ")"
    End If

    openedByUs = True
    Set GetOrOpenWorkbook = wb
End Function

Private Function FindTransactionSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim sName As String
    Dim preferred As Variant, p As Long

    ' Priority 1: known sheet names that also carry a Transaction ID header
    preferred = Array("lookback", "raw transactions", "inscope", "in scope", "in-scope")
    For p = 0 To UBound(preferred)
        For Each ws In wb.Worksheets
            If InStr(1, ws.Name, CStr(preferred(p)), vbTextCompare) > 0 Then
                If FindHeaderRow(ws, FLD_TXID) > 0 Then
                    Set FindTransactionSheet = ws
                    Exit Function
                End If
            End If
        Next ws
    Next p

    ' Priority 2: first other sheet with a Transaction ID header
    For Each ws In wb.Worksheets
        sName = LCase$(ws.Name)
        If InStr(sName, "pivot") = 0 And InStr(sName, "dashboard") = 0 And InStr(sName, "discrepancy") = 0 Then
            If FindHeaderRow(ws, FLD_TXID) > 0 Then
                Set FindTransactionSheet = ws
                Exit Function
            End If
        End If
    Next ws

    Set FindTransactionSheet = Nothing
End Function

Private Function Pick_Folder(ByVal promptTitle As String) As String
    Dim result As String
#If Mac Then
    Dim script As String, errNum As Long
    script = "return POSIX path of (choose folder with prompt """ & Replace(promptTitle, """", "'") & """)"
    On Error Resume Next
    result = MacScript(script)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then
        ' MacScript also raises when Cancel is pressed; offer a manual path in case the picker is blocked
        result = ""
        If MsgBox("No folder was selected." & vbCrLf & vbCrLf & _
                  "If the folder picker did not appear, click Yes to paste the folder path instead.", _
                  vbQuestion + vbYesNo, "Select Folder") = vbYes Then
            result = InputBox(promptTitle & vbCrLf & vbCrLf & "Example: /Users/you/Documents/Lookback/Old", "Folder Path")
        End If
    End If
#Else
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    fd.Title = promptTitle
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then result = fd.SelectedItems(1)
#End If
    result = Trim$(result)
    If result <> "" Then
        If Right$(result, 1) <> Application.PathSeparator Then result = result & Application.PathSeparator
    End If
    Pick_Folder = result
End Function

Private Function Pick_Excel_File(ByVal promptTitle As String) As String
    Dim vFile As Variant
#If Mac Then
    ' File filters are not supported on Mac
    vFile = Application.GetOpenFilename()
#Else
    vFile = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb", , promptTitle)
#End If
    If VarType(vFile) = vbBoolean Then
        Pick_Excel_File = ""
    ElseIf Not IsExcelFileName(CStr(vFile)) Then
        MsgBox "Please select an Excel file (.xlsx, .xlsm, .xls or .xlsb).", vbExclamation, "Not an Excel File"
        Pick_Excel_File = ""
    Else
        Pick_Excel_File = CStr(vFile)
    End If
End Function

Private Sub GrantMacAccess(ByVal paths As Variant)
#If Mac Then
    #If MAC_OFFICE_VERSION >= 15 Then
    ' Excel 2016+ for Mac is sandboxed: ask once for access to the chosen files/folders
    Dim granted As Boolean
    On Error Resume Next
    granted = GrantAccessToMultipleFiles(paths)
    On Error GoTo 0
    #End If
#End If
End Sub

Private Function ListExcelFiles(ByVal folderPath As String, ByVal excludePath As String) As Collection
    Dim col As New Collection
    Dim names As New Collection
    Dim sFile As String
    Dim arr As Variant, i As Long

    If Right$(folderPath, 1) <> Application.PathSeparator Then folderPath = folderPath & Application.PathSeparator

    ' Dir without a wildcard works on both Windows and Mac; filter by extension here
    sFile = Dir(folderPath, vbReadOnly)
    Do While sFile <> ""
        If IsExcelFileName(sFile) And Left$(sFile, 2) <> "~$" And Left$(sFile, 2) <> "._" _
           And InStr(1, sFile, "AML_Period_Comparison", vbTextCompare) = 0 _
           And StrComp(folderPath & sFile, excludePath, vbTextCompare) <> 0 Then
            names.Add sFile
        End If
        sFile = Dir()
    Loop

    arr = CollectionToArray(names)
    SortStrings arr
    For i = LBound(arr) To UBound(arr)
        col.Add folderPath & arr(i)
    Next i
    Set ListExcelFiles = col
End Function

Private Function IsExcelFileName(ByVal fName As String) As Boolean
    Dim ext As String, p As Long
    p = InStrRev(fName, ".")
    If p = 0 Then Exit Function
    ext = LCase$(Mid$(fName, p + 1))
    IsExcelFileName = (ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "xlsb")
End Function

Private Function FolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(p) And vbDirectory) = vbDirectory)
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(p) And vbDirectory) = 0)
End Function

Private Function StripExtension(ByVal fName As String) As String
    Dim p As Long
    p = InStrRev(fName, ".")
    If p > 0 Then
        StripExtension = Left$(fName, p - 1)
    Else
        StripExtension = fName
    End If
End Function

Private Function StripExcelExtension(ByVal fPath As String) As String
    Dim p As Long, s As Long, ext As String
    p = InStrRev(fPath, ".")
    s = InStrRev(fPath, Application.PathSeparator)
    StripExcelExtension = fPath
    If p > s And p > 0 Then
        ext = LCase$(Mid$(fPath, p + 1))
        If ext = "xls" Or ext = "xlsm" Or ext = "xlsb" Or ext = "csv" Then StripExcelExtension = Left$(fPath, p - 1)
    End If
End Function

Private Function GetFileName(ByVal fPath As String) As String
    Dim pos As Long
    pos = InStrRev(fPath, Application.PathSeparator)
    If pos = 0 Then pos = InStrRev(fPath, "/")
    If pos > 0 Then
        GetFileName = Mid$(fPath, pos + 1)
    Else
        GetFileName = fPath
    End If
End Function

' Derives ECM ID, Alert ID and a match key from a file name, e.g.
'   138896_ALERT2365101_Inscope transactions 06.02.2025 to 06.03.2026.xlsx
'   -> ECM "138896", Alert "ALERT2365101", key "138896_ALERT2365101"
' Also handles "Alert_2365101", "ALERT-2365101", "ALERT 2365101".
Private Sub ExtractFileIDs(ByVal fPath As String, ByRef ecmID As String, ByRef alertID As String, ByRef matchKey As String)
    Dim baseName As String, u As String
    Dim p As Long, j As Long, k As Long, startPos As Long, alertPos As Long
    Dim alertNum As String, prefix As String, lead As String
    Dim tokens() As String, t As Long, firstNum As String, secondNum As String

    baseName = Trim$(StripExtension(GetFileName(fPath)))
    u = UCase$(baseName)
    ecmID = "-"
    alertID = ""
    matchKey = ""

    ' Find "ALERT" followed (after optional separators) by a digit
    startPos = 1
    Do
        p = InStr(startPos, u, "ALERT")
        If p = 0 Then Exit Do
        j = p + 5
        Do While j <= Len(u)
            If InStr("_- #:", Mid$(u, j, 1)) = 0 Then Exit Do
            j = j + 1
        Loop
        If j <= Len(u) Then
            If Mid$(u, j, 1) Like "#" Then
                k = j
                Do While k <= Len(u)
                    If Not (Mid$(u, k, 1) Like "[0-9A-Z]") Then Exit Do
                    k = k + 1
                Loop
                alertNum = Mid$(u, j, k - j)
                alertPos = p
                Exit Do
            End If
        End If
        startPos = p + 5
    Loop

    If alertNum <> "" Then
        alertID = "ALERT" & alertNum
        prefix = TrimSeparators(Left$(baseName, alertPos - 1))
        If prefix <> "" Then
            lead = LastToken(prefix)
            If lead Like "*#*" Then ecmID = lead
        End If
        If ecmID <> "-" Then
            matchKey = UCase$(ecmID) & "_" & alertID
        Else
            matchKey = alertID
        End If
    Else
        ' No ALERT token: use up to two leading numeric tokens, e.g. "138896_2365101_..."
        tokens = Split(Application.WorksheetFunction.Trim(Replace(Replace(baseName, "_", " "), "-", " ")), " ")
        For t = 0 To UBound(tokens)
            If Not IsAllDigits(tokens(t)) Then Exit For
            If firstNum = "" Then
                firstNum = tokens(t)
            Else
                secondNum = tokens(t)
                Exit For
            End If
        Next t
        If secondNum <> "" Then
            ecmID = firstNum
            alertID = secondNum
            matchKey = firstNum & "_" & secondNum
        ElseIf firstNum <> "" Then
            alertID = firstNum
            matchKey = firstNum
        Else
            alertID = baseName
            matchKey = UCase$(baseName)
        End If
    End If
End Sub

Private Function TrimSeparators(ByVal s As String) As String
    Do While Len(s) > 0
        If InStr("_- ", Right$(s, 1)) = 0 Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    Do While Len(s) > 0
        If InStr("_- ", Left$(s, 1)) = 0 Then Exit Do
        s = Mid$(s, 2)
    Loop
    TrimSeparators = s
End Function

Private Function LastToken(ByVal s As String) As String
    Dim i As Long
    For i = Len(s) To 1 Step -1
        If InStr("_- ", Mid$(s, i, 1)) > 0 Then
            LastToken = Mid$(s, i + 1)
            Exit Function
        End If
    Next i
    LastToken = s
End Function

' Date range found in the data, plus the period written in the file name (if any)
Private Function FormatPeriod(ByVal dMin As Variant, ByVal dMax As Variant, ByVal fPath As String) As String
    Dim s As String, fp As String
    If IsDate(dMin) And IsDate(dMax) Then s = Format(dMin, "yyyy-mm-dd") & " to " & Format(dMax, "yyyy-mm-dd")
    fp = FilePeriodText(fPath)
    If fp <> "" Then
        If s = "" Then
            s = "file: " & fp
        Else
            s = s & "  |  file: " & fp
        End If
    End If
    If s = "" Then s = "N/A"
    FormatPeriod = s
End Function

' Finds date-like tokens in a file name, e.g. "06.02.2025 to 06.03.2026"
Private Function FilePeriodText(ByVal fPath As String) As String
    Dim baseName As String, tok As String
    Dim i As Long, j As Long
    Dim found1 As String, found2 As String

    baseName = StripExtension(GetFileName(fPath))
    i = 1
    Do While i <= Len(baseName)
        If Mid$(baseName, i, 1) Like "#" Then
            j = i
            Do While j <= Len(baseName)
                If Not (Mid$(baseName, j, 1) Like "[0-9./-]") Then Exit Do
                j = j + 1
            Loop
            tok = Mid$(baseName, i, j - i)
            Do While Len(tok) > 0
                If InStr("./-", Right$(tok, 1)) = 0 Then Exit Do
                tok = Left$(tok, Len(tok) - 1)
            Loop
            If LooksLikeDateToken(tok) Then
                If found1 = "" Then
                    found1 = tok
                ElseIf found2 = "" Then
                    found2 = tok
                End If
            End If
            i = j
        Else
            i = i + 1
        End If
    Loop

    If found1 <> "" And found2 <> "" Then
        FilePeriodText = found1 & " to " & found2
    Else
        FilePeriodText = found1
    End If
End Function

Private Function LooksLikeDateToken(ByVal tok As String) As Boolean
    Dim sep As String, parts() As String
    If InStr(tok, ".") > 0 Then
        sep = "."
    ElseIf InStr(tok, "/") > 0 Then
        sep = "/"
    ElseIf InStr(tok, "-") > 0 Then
        sep = "-"
    Else
        Exit Function
    End If
    parts = Split(tok, sep)
    If UBound(parts) <> 2 Then Exit Function
    If Not (IsAllDigits(parts(0)) And IsAllDigits(parts(1)) And IsAllDigits(parts(2))) Then Exit Function
    If Len(parts(0)) = 4 Then
        LooksLikeDateToken = (Len(parts(1)) <= 2 And Len(parts(2)) <= 2)
    Else
        LooksLikeDateToken = (Len(parts(0)) <= 2 And Len(parts(1)) <= 2 And (Len(parts(2)) = 2 Or Len(parts(2)) = 4))
    End If
End Function

Private Function FindSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set FindSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetOrCreateWorksheet(ByVal wb As Workbook, ByVal sheetName As String, Optional ByVal isFirst As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Set ws = FindSheet(wb, sheetName)

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

Private Sub SaveAndSpeedUp(ByRef st As AppState)
    st.ScreenUpdating = Application.ScreenUpdating
    st.DisplayAlerts = Application.DisplayAlerts
    st.EnableEvents = Application.EnableEvents
    On Error Resume Next
    st.Calculation = Application.Calculation
    st.HasCalculation = (Err.Number = 0)
    Err.Clear
    Application.Calculation = xlCalculationManual
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
End Sub

Private Sub RestoreApp(ByRef st As AppState)
    On Error Resume Next
    If st.HasCalculation Then Application.Calculation = st.Calculation
    Application.ScreenUpdating = st.ScreenUpdating
    Application.DisplayAlerts = st.DisplayAlerts
    Application.EnableEvents = st.EnableEvents
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' =========================================================================
' [HELPER] 12. SHEET READING AND HEADER DETECTION
' =========================================================================
' Always returns a 1-based 2D array, even for a single cell
Private Function ReadBlock(ByVal ws As Worksheet, ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long) As Variant
    Dim v As Variant
    Dim arr() As Variant
    v = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value2
    If IsArray(v) Then
        ReadBlock = v
    Else
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = v
        ReadBlock = arr
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not f Is Nothing Then LastUsedRow = f.Row
End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not f Is Nothing Then LastUsedCol = f.Column
End Function

' Returns the first row (within HEADER_SCAN_ROWS) holding a header for fieldId, or 0
Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal fieldId As Long) As Long
    Dim lastRow As Long, lastCol As Long, maxR As Long, r As Long
    Dim block As Variant

    lastCol = LastUsedCol(ws)
    lastRow = LastUsedRow(ws)
    If lastCol = 0 Or lastRow = 0 Then Exit Function
    maxR = HEADER_SCAN_ROWS
    If lastRow < maxR Then maxR = lastRow

    block = ReadBlock(ws, 1, 1, maxR, lastCol)
    For r = 1 To maxR
        If GetFieldCol(block, r, lastCol, fieldId) > 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub DetectColumns(ByVal hdr As Variant, ByVal lastCol As Long, ByRef cm As ColMap)
    cm.TxID = GetFieldCol(hdr, 1, lastCol, FLD_TXID)
    cm.Alert = GetFieldCol(hdr, 1, lastCol, FLD_ALERT)
    cm.TxDate = GetFieldCol(hdr, 1, lastCol, FLD_DATE)
    cm.Amount = GetFieldCol(hdr, 1, lastCol, FLD_AMOUNT)
    cm.Account = GetFieldCol(hdr, 1, lastCol, FLD_ACCOUNT)
    cm.Desc = GetFieldCol(hdr, 1, lastCol, FLD_DESC)
    cm.CP = GetFieldCol(hdr, 1, lastCol, FLD_CP)
    cm.DrCr = GetFieldCol(hdr, 1, lastCol, FLD_DRCR)
    cm.Ccy = GetFieldCol(hdr, 1, lastCol, FLD_CCY)
    cm.OrigName = GetFieldCol(hdr, 1, lastCol, FLD_ORIG)
    cm.BenName = GetFieldCol(hdr, 1, lastCol, FLD_BEN)
End Sub

' Best-scoring column for a field in row r of block (exact names beat partial matches;
' earlier names in the list beat later ones; ties go to the leftmost column)
Private Function GetFieldCol(ByVal block As Variant, ByVal r As Long, ByVal lastCol As Long, ByVal fieldId As Long) As Long
    Dim c As Long, s As Long, bestScore As Long
    Dim h As String
    For c = 1 To lastCol
        h = NormText(CellText(block(r, c)))
        If h <> "" Then
            s = HeaderScore(h, fieldId)
            If s > bestScore Then
                bestScore = s
                GetFieldCol = c
            End If
        End If
    Next c
End Function

Private Function HeaderScore(ByVal h As String, ByVal fieldId As Long) As Long
    Dim exacts As Variant, partials As Variant, excludes As Variant
    Dim i As Long

    FieldSpec fieldId, exacts, partials, excludes

    For i = 0 To UBound(exacts)
        If h = exacts(i) Then
            HeaderScore = 1000 - i
            Exit Function
        End If
    Next i

    For i = 0 To UBound(excludes)
        If InStr(" " & h & " ", " " & excludes(i) & " ") > 0 Then Exit Function
    Next i

    For i = 0 To UBound(partials)
        If InStr(h, partials(i)) > 0 Then
            HeaderScore = 500 - i
            Exit Function
        End If
    Next i
End Function

' Header names are compared after NormText (lower case, punctuation -> space).
' exacts: whole-header matches in priority order. partials: substrings.
' excludes: whole words that disqualify a partial match.
Private Sub FieldSpec(ByVal fieldId As Long, ByRef exacts As Variant, ByRef partials As Variant, ByRef excludes As Variant)
    Select Case fieldId
        Case FLD_TXID
            exacts = Array("transaction id", "txn id", "trans id", "transactionid", "txnid", "transaction number", "transaction no", "txn no")
            partials = Array("transaction id", "txn id")
            excludes = Array("velocity", "alert", "related", "original", "parent", "reversal", "linked")
        Case FLD_ALERT
            exacts = Array("is alerted transaction", "is alerted", "alerted", "alerted transaction", "is alerted txn", "alerted flag", "alert flag", "is alerted flag")
            partials = Array("is alerted", "alerted")
            excludes = Array("id", "information", "info", "type", "date", "amount", "count", "reason", "score")
        Case FLD_DATE
            exacts = Array("transaction date", "txn date", "trans date", "posting date", "post date", "value date", "booking date", "date")
            partials = Array("transaction date", "txn date", "date")
            excludes = Array("update", "updated", "birth", "dob", "open", "opened", "created", "modified", "alert", "load", "run", "report")
        Case FLD_AMOUNT
            exacts = Array("transaction amount", "txn amount", "trans amount", "amount", "amount usd", "base amount", "usd amount")
            partials = Array("transaction amount", "txn amount", "amount")
            excludes = Array("count", "sum", "total", "local", "limit", "threshold")
        Case FLD_ACCOUNT
            exacts = Array("account no", "account number", "account", "account id", "customer account", "acct no", "account num")
            partials = Array("account")
            excludes = Array("counterparty", "beneficiary", "originator", "party", "bank", "type", "name", "status")
        Case FLD_DESC
            exacts = Array("transaction description", "transaction code description", "description", "txn description", "narrative", "remarks", "memo")
            partials = Array("description", "desc", "narrative")
            excludes = Array("alert")
        Case FLD_CP
            exacts = Array("counterparty", "counterparty name", "counter party", "counter party name", "cp name")
            partials = Array("counterparty name", "counter party name", "counterparty", "counter party")
            excludes = Array("account", "acct", "country", "id", "address", "type", "bank", "number", "no", "city", "state")
        Case FLD_DRCR
            exacts = Array("dr cr", "drcr", "cr dr", "crdr", "debit credit", "credit debit", "dr cr indicator", "dr cr flag", "debit credit indicator", "credit debit indicator", "direction")
            partials = Array("dr cr", "debit credit", "credit debit")
            excludes = Array("amount")
        Case FLD_CCY
            exacts = Array("currency", "ccy", "currency code", "transaction currency", "txn currency")
            partials = Array("currency")
            excludes = Array("amount", "local", "base")
        Case FLD_ORIG
            exacts = Array("originator name", "orig name", "ordering customer name", "remitter name", "sender name")
            partials = Array("originator name")
            excludes = Array("first", "last", "middle")
        Case FLD_BEN
            exacts = Array("beneficiary name", "ben name", "bene name", "receiver name", "beneficiary customer name")
            partials = Array("beneficiary name")
            excludes = Array("first", "last", "middle")
        Case Else
            exacts = Array()
            partials = Array()
            excludes = Array()
    End Select
End Sub

' =========================================================================
' [HELPER] 13. VALUE PARSING
' =========================================================================
' Cell value -> clean text (never raises on error cells)
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger, vbByte
            If v = Int(v) And Abs(v) < 1E+15 Then
                CellText = Format$(v, "0")
            Else
                CellText = CStr(v)
            End If
        Case vbBoolean
            If v Then
                CellText = "TRUE"
            Else
                CellText = "FALSE"
            End If
        Case Else
            CellText = Trim$(Replace(CStr(v), Chr(160), " "))
    End Select
End Function

' Matching key for a transaction ID: upper case; numeric IDs lose leading zeros
' so "0012345" (text) matches 12345 (number)
Private Function TxKey(ByVal s As String) As String
    s = UCase$(Trim$(s))
    If IsAllDigits(s) Then
        Do While Len(s) > 1 And Left$(s, 1) = "0"
            s = Mid$(s, 2)
        Loop
    End If
    TxKey = s
End Function

Private Function IsYes(ByVal s As String) As Boolean
    Select Case UCase$(Trim$(s))
        Case "YES", "Y", "TRUE", "1", "ALERTED"
            IsYes = True
    End Select
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    IsAllDigits = Not (s Like "*[!0-9]*")
End Function

' Lower case, every non-alphanumeric character -> single space, trimmed
Private Function NormText(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, lastSpace As Boolean
    s = LCase$(s)
    lastSpace = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9]" Then
            out = out & ch
            lastSpace = False
        ElseIf Not lastSpace Then
            out = out & " "
            lastSpace = True
        End If
    Next i
    NormText = Trim$(out)
End Function

' Excel serial / real date / text date -> Date. Text dates use TEXT_DATE_ORDER,
' not the PC's regional settings. ok = False when the value is blank or unreadable.
Private Function ParseDateValue(ByVal v As Variant, ByRef ok As Boolean) As Date
    Dim s As String, sep As String, parts() As String
    Dim a As Long, b As Long, c As Long, y As Long, m As Long, d As Long, tmp As Long
    Dim p As Long, dbl As Double, result As Date

    ok = False
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function

    Select Case VarType(v)
        Case vbDate
            ParseDateValue = DateSerial(Year(v), Month(v), Day(v))
            ok = True
            Exit Function
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger, vbByte
            dbl = CDbl(v)
            If dbl >= 18264 And dbl < 73051 Then        ' 1950-01-01 .. 2099-12-31
                ParseDateValue = CDate(Int(dbl))
                ok = True
            End If
            Exit Function
        Case vbString
            ' handled below
        Case Else
            Exit Function
    End Select

    s = Trim$(Replace(CStr(v), Chr(160), " "))
    If s = "" Then Exit Function

    ' Drop any time part: "2025-10-31T00:00:00", "10/31/2025 14:05"
    If Len(s) > 10 And Mid$(s, 11, 1) = "T" Then s = Left$(s, 10)
    p = InStr(s, " ")
    If p > 0 Then s = Left$(s, p - 1)

    If InStr(s, "/") > 0 Then
        sep = "/"
    ElseIf InStr(s, "-") > 0 Then
        sep = "-"
    ElseIf InStr(s, ".") > 0 Then
        sep = "."
    End If

    If sep <> "" Then
        parts = Split(s, sep)
        If UBound(parts) = 2 Then
            If IsAllDigits(parts(0)) And IsAllDigits(parts(1)) And IsAllDigits(parts(2)) Then
                If Len(parts(0)) > 4 Or Len(parts(1)) > 2 Or Len(parts(2)) > 4 Then Exit Function
                a = CLng(parts(0))
                b = CLng(parts(1))
                c = CLng(parts(2))
                If Len(parts(0)) = 4 Then
                    y = a
                    m = b
                    d = c
                ElseIf TEXT_DATE_ORDER = "DMY" Then
                    d = a
                    m = b
                    y = c
                Else
                    m = a
                    d = b
                    y = c
                End If
                If m > 12 And d <= 12 Then
                    tmp = m
                    m = d
                    d = tmp
                End If
                If y < 100 Then
                    If y < 50 Then
                        y = y + 2000
                    Else
                        y = y + 1900
                    End If
                End If
                If BuildDate(y, m, d, result) Then
                    ParseDateValue = result
                    ok = True
                End If
                Exit Function
            End If
        End If
    End If

    ' Numeric text holding an Excel serial
    If IsAllDigits(s) Then
        If Len(s) <= 5 Then
            dbl = CDbl(s)
            If dbl >= 18264 And dbl < 73051 Then
                ParseDateValue = CDate(Int(dbl))
                ok = True
            End If
        End If
        Exit Function
    End If

    ' Month-name formats ("31-Oct-2025", "Oct 31 2025") are unambiguous
    If s Like "*[A-Za-z]*" Then
        If IsDate(CStr(v)) Then
            result = CDate(CStr(v))
            ParseDateValue = DateSerial(Year(result), Month(result), Day(result))
            ok = True
        End If
    End If
End Function

Private Function BuildDate(ByVal y As Long, ByVal m As Long, ByVal d As Long, ByRef result As Date) As Boolean
    If y < 1900 Or y > 2100 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    result = DateSerial(y, m, d)
    BuildDate = (Day(result) = d And Month(result) = m)
End Function

' Numbers and numeric text such as "1,234.50", "$1,234.50", "(1,234.50)"
Private Function ParseAmount(ByVal v As Variant, ByRef ok As Boolean) As Double
    Dim s As String, neg As Boolean, amt As Double
    ok = False
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function

    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger, vbByte
            ParseAmount = CDbl(v)
            ok = True
        Case vbString
            s = Trim$(Replace(CStr(v), Chr(160), ""))
            If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
                neg = True
                s = Mid$(s, 2, Len(s) - 2)
            End If
            s = Replace(Replace(Replace(s, ",", ""), "$", ""), " ", "")
            If s Like "*#*" And Not (s Like "*[!0-9.+-]*") Then
                amt = Val(s)
                If neg Then amt = -amt
                ParseAmount = amt
                ok = True
            End If
    End Select
End Function

' Display text for audit values
Private Function ValText(ByVal v As Variant) As String
    Dim s As String
    If IsEmpty(v) Then
        s = "(blank)"
    ElseIf VarType(v) = vbDate Then
        s = Format(v, "yyyy-mm-dd")
    ElseIf VarType(v) = vbDouble Then
        s = Format(v, "#,##0.00")
    Else
        s = CStr(v)
        If s = "" Then s = "(blank)"
    End If
    ValText = s
End Function

Private Sub AddListItem(ByRef lst As String, ByVal item As String)
    If lst = "" Then
        lst = item
    Else
        lst = lst & ", " & item
    End If
End Sub

' Union of two comma-separated lists, sorted
Private Function MergeList(ByVal a As String, ByVal b As String) As String
    Dim d As Object, v As Variant, s As String
    Set d = CreateLookupDict()
    For Each v In Split(a & "," & b, ",")
        s = Trim$(CStr(v))
        If s <> "" Then
            If Not DictExists(d, s) Then DictAdd d, s, s
        End If
    Next v
    MergeList = JoinSortedKeys(d)
End Function

Private Function JoinSortedKeys(ByVal d As Object) As String
    Dim arr As Variant
    arr = DictKeys(d)
    If UBound(arr) < LBound(arr) Then Exit Function
    SortStrings arr
    JoinSortedKeys = Join(arr, ", ")
End Function

Private Function CollectionToArray(ByVal col As Collection) As Variant
    Dim arr() As String
    Dim i As Long, v As Variant
    If col.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    ReDim arr(0 To col.Count - 1)
    For Each v In col
        arr(i) = CStr(v)
        i = i + 1
    Next v
    CollectionToArray = arr
End Function

' Insertion sort, case-insensitive (lists here are small)
Private Sub SortStrings(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    If UBound(arr) <= LBound(arr) Then Exit Sub
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = arr(i)
        j = i - 1
        Do While j >= LBound(arr)
            If StrComp(CStr(arr(j)), CStr(tmp), vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub

' =========================================================================
' [DICT] 14. KEY-VALUE LOOKUP (Scripting.Dictionary on Windows,
'            Collection fallback on Mac). Keys are case-insensitive on both.
' =========================================================================
Private Function CreateLookupDict() As Object
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo 0

    If d Is Nothing Then
        Set d = New Collection
    Else
        d.CompareMode = 1 ' vbTextCompare - same behaviour as Collection keys
    End If
    Set CreateLookupDict = d
End Function

Private Function DictExists(ByVal d As Object, ByVal k As String) As Boolean
    Dim dummy As Variant
    If TypeName(d) = "Dictionary" Then
        DictExists = d.Exists(k)
    Else
        On Error Resume Next
        dummy = d(k)
        DictExists = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If
End Function

Private Sub DictAdd(ByVal d As Object, ByVal k As String, ByVal v As Variant)
    If TypeName(d) = "Dictionary" Then
        d.Add k, v
    Else
        d.Add Array(k, v), k      ' Collection items keep their key at index 0
    End If
End Sub

Private Function DictGet(ByVal d As Object, ByVal k As String) As Variant
    Dim pair As Variant
    If TypeName(d) = "Dictionary" Then
        DictGet = d(k)
    Else
        pair = d(k)
        DictGet = pair(1)
    End If
End Function

Private Function DictCount(ByVal d As Object) As Long
    DictCount = d.Count
End Function

Private Function DictKeys(ByVal d As Object) As Variant
    Dim keysArr() As String
    Dim i As Long, pair As Variant
    If TypeName(d) = "Dictionary" Then
        DictKeys = d.Keys
    Else
        If d.Count = 0 Then
            DictKeys = Array()
            Exit Function
        End If
        ReDim keysArr(0 To d.Count - 1)
        For Each pair In d
            keysArr(i) = CStr(pair(0))
            i = i + 1
        Next pair
        DictKeys = keysArr
    End If
End Function
