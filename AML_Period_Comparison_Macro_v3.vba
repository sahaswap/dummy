Attribute VB_Name = "AML_Period_Comparison"
' =========================================================================
' [AML] AML TRANSACTION MONITORING: LOOKBACK PERIOD BATCH COMPARISON COCKPIT
' =========================================================================
' Version: 3.4
'
' Purpose:
'  1. Compares the Old vs. New transaction file of each alert when the review
'     lookback dates change, and answers:
'       - Does the New file contain new transactions? How many, how much?
'       - Do the new transactions involve counterparties that do not appear
'         anywhere in the Old file (newly identified counterparties)?
'       - Were any transactions dropped?
'  2. Dashboard with KPI cards, one-click buttons and a status per alert.
'  3. Sheet "New_CP_Pivots": a summary per alert plus one pivot table per
'     alert listing its newly identified counterparties.
'  4. Optional: updates each alert's escalation narrative (ECMID_ALERTID_*.docx)
'     with the New file figures - opening sentence, Total Suspicious Dollar
'     Amount, Date Range of Suspicious Activity, the "Between ..." line and the
'     first sentence of the "Specifically, between ..." line. Edited copies go
'     to a new subfolder; each changed value is highlighted yellow. Originals
'     are never changed. Needs Microsoft Word.
'  5. Single-pair and batch-folder comparison. Windows and Mac (Excel 2016+).
'
' Counterparty rules:
'  - Counterparty comes from a "Counterparty" column. If a sheet has none, it
'    is taken from Originator Name (CR) or Beneficiary Name (DR).
'  - A counterparty is "known" if it appears anywhere in the Old file: the
'    transaction sheet, any other sheet with a Counterparty column, or any
'    pivot table field named Counterparty.
'  - Names are matched ignoring case, punctuation, extra spaces and SWIFT
'    line tags, so "1/ABC LTD." and "ABC LTD" are the same counterparty.
'
' Settings: see "User Settings" below.
' =========================================================================
Option Explicit

' --- User Settings ---
' How text dates such as "10/07/2025" are read. "MDY" = US (month first),
' "DMY" = day first. Real Excel dates are never affected by this setting.
Private Const TEXT_DATE_ORDER As String = "MDY"
' How many rows from the top of a sheet are searched for the header row.
Private Const HEADER_SCAN_ROWS As Long = 10
' Narrative update: add Word comments (Old vs New summary, new counterparties,
' rule not mentioned). Lines that cannot be found are always commented.
Private Const NARRATIVE_ADD_COMMENTS As Boolean = True

' --- Design System Palette (Tailored Slate Theme) ---
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
Private Const SHEET_NEWCP_PIVOTS As String = "New_CP_Pivots"
Private Const SHEET_LEGACY_MODIFIED As String = "Discrepancy_Modified_Txns"   ' created by v3.0; removed on run

' --- Dashboard Layout ---
Private Const DASH_VERSION As String = "AML-PERIOD-COMPARISON-3.2"
Private Const DASH_TITLE As String = "AML TRANSACTION MONITORING - LOOKBACK PERIOD BATCH RECONCILIATION COCKPIT"
Private Const FIRST_DATA_ROW As Long = 15
Private Const DASH_COLS As Long = 18               ' columns B:S

' --- Audit Sheet Kinds ---
Private Const AUDIT_ADDED As Long = 1
Private Const AUDIT_DROPPED As Long = 2

' --- "New Counterparty?" values ---
Private Const CP_YES As String = "Yes"
Private Const CP_NO As String = "No"

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
Private Const FLD_RULE As Long = 12

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
    Rule As Long
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
    AlertAmt As Double         ' alerted transactions: gross amount
    AlertMin As Variant        ' alerted transactions: first date
    AlertMax As Variant        ' alerted transactions: last date
    RuleText As String         ' rule(s) named on the alerted rows, e.g. "Many to one fund transfer - Wires"
    CrCount As Long            ' unique credit transactions
    CrAmt As Double
    DrCount As Long            ' unique debit transactions
    DrAmt As Double
End Type

Private Type PairResult
    AddedCount As Long
    AddedAlerted As Long
    AddedAmount As Double      ' sum of absolute amounts of the new transactions
    DroppedCount As Long
    DroppedAlerted As Long
    NewCPTxns As Long          ' new transactions with a new counterparty
    NewCPDistinct As Long      ' distinct new counterparties
    NewCPNames As String       ' "NAME A; NAME B"
    NewCPAgg As Object         ' normalised name -> Array(name, txn count, amount)
    CPUnknown As Long          ' new transactions whose counterparty could not be checked
    AddedDetail As String      ' one line per new transaction (first 25)
    NewCPDetail As String      ' "NAME (n txn(s), $x); ..."
End Type

Private Type BatchTotals
    Evaluated As Long
    Unmatched As Long
    DupFiles As Long
    Errors As Long
    PairsWithNewTxns As Long
    PairsWithNewCP As Long
    Added As Long
    AddedAlerted As Long
    AddedAmount As Double
    Dropped As Long
    NewCP As Long
    Currencies As String
    AlertRows As Collection    ' Array(ecm, alert, new txns, amount, new CPs, names, aggregate)
    NarrUpdated As Long
    NarrFlagged As Long
    NarrNoChange As Long
    NarrNotFound As Long
    NarrErrors As Long
    NarrOutDir As String
End Type

Private Type NarrCtx
    EcmID As String
    AlertID As String
    OldFile As String
    NewFile As String
    OldCount As Long
    NewCount As Long
    OldTotal As Double
    NewTotal As Double
    OldMin As Variant
    OldMax As Variant
    NewMin As Variant
    NewMax As Variant
    OldAlertCount As Long
    NewAlertCount As Long
    OldAlertAmt As Double
    NewAlertAmt As Double
    OldAlertMin As Variant
    OldAlertMax As Variant
    NewAlertMin As Variant
    NewAlertMax As Variant
    HasDrCr As Boolean
    OldCrCount As Long
    OldCrAmt As Double
    OldDrCount As Long
    OldDrAmt As Double
    NewCrCount As Long
    NewCrAmt As Double
    NewDrCount As Long
    NewDrAmt As Double
    OldPeriodFrom As Variant
    OldPeriodTo As Variant
    NewPeriodFrom As Variant
    NewPeriodTo As Variant
    RuleText As String
    AddedCount As Long
    AddedAmount As Double
    AddedAlerted As Long
    AddedDetail As String
    DroppedCount As Long
    NewCPDistinct As Long
    NewCPDetail As String
End Type

Private Type AppState
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    EnableEvents As Boolean
    Calculation As Long
    HasCalculation As Boolean
End Type

' Word instance used for narrative updates (late bound)
Private mWord As Object
Private mWordCreated As Boolean
Private mWordAlerts As Long

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

    RunJobs wbDashboard, jobs, "Batch", AskNarrativeFolder()
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
    RunJobs wbDashboard, jobs, "Single Alert", AskNarrativeFolder()
End Sub

' =========================================================================
' [RESET] 3. USER-FACING MACRO: RESET / CLEAR DASHBOARD
' =========================================================================
Public Sub Clear_Comparison_Dashboard()
    Dim wb As Workbook
    Dim sheetName As Variant
    Dim ws As Worksheet

    If MsgBox("Clear the Comparison Dashboard and delete the discrepancy and pivot sheets?", _
              vbQuestion + vbYesNo, "Confirm Reset") <> vbYes Then Exit Sub

    Set wb = ActiveWorkbook
    Setup_Comparison_Dashboard wb, True

    Application.DisplayAlerts = False
    For Each sheetName In Array(SHEET_ADDED_TXNS, SHEET_DROPPED_TXNS, SHEET_NEWCP_PIVOTS, SHEET_LEGACY_MODIFIED)
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
    Dim wsDash As Worksheet
    Dim defaultFileName As String, savePath As Variant, finalPath As String
    Dim sheetName As Variant, sheetList() As Variant, n As Long
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

    ' Sheets are copied together so the pivots keep pointing at the copied data
    ReDim sheetList(0 To 3)
    For Each sheetName In Array(SHEET_DASHBOARD, SHEET_ADDED_TXNS, SHEET_DROPPED_TXNS, SHEET_NEWCP_PIVOTS)
        If Not FindSheet(wbSource, CStr(sheetName)) Is Nothing Then
            sheetList(n) = CStr(sheetName)
            n = n + 1
        End If
    Next sheetName
    ReDim Preserve sheetList(0 To n - 1)

    SaveAndSpeedUp st
    On Error GoTo ExportFail

    wbSource.Worksheets(sheetList).Copy
    Set wbNew = ActiveWorkbook
    DeleteButtonShapes wbNew.Worksheets(SHEET_DASHBOARD)

    wbNew.Worksheets(SHEET_DASHBOARD).Activate
    wbNew.SaveAs Filename:=finalPath, FileFormat:=51     ' 51 = xlOpenXMLWorkbook (.xlsx)

    RestoreApp st
    MsgBox "AML Lookback Report successfully exported to:" & vbCrLf & vbCrLf & _
           finalPath, vbInformation, "Export Successful"
    Exit Sub

ExportFail:
    errDesc = Err.Description
    Resume ExportCleanup

ExportCleanup:
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

Public Sub View_New_CP_Pivots()
    ShowSheet SHEET_NEWCP_PIVOTS
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
Private Sub RunJobs(ByVal wbDash As Workbook, ByVal jobs As Collection, ByVal modeLabel As String, _
                    ByVal narrativeDir As String)
    Dim wsDash As Worksheet, wsAdded As Worksheet, wsDropped As Worksheet, ws As Worksheet
    Dim st As AppState
    Dim tot As BatchTotals
    Dim job As Variant
    Dim curRow As Long, jobNo As Long
    Dim stepName As String, summary As String

    SaveAndSpeedUp st
    On Error GoTo ErrorHandler
    Set tot.AlertRows = New Collection

    stepName = "Setting Up Worksheets"
    Set wsDash = Setup_Comparison_Dashboard(wbDash)
    Set ws = FindSheet(wbDash, SHEET_NEWCP_PIVOTS)
    If Not ws Is Nothing Then ClearPivotSheet ws
    Set ws = FindSheet(wbDash, SHEET_LEGACY_MODIFIED)      ' sheet from v3.0, no longer used
    If Not ws Is Nothing Then ws.Delete
    Set wsAdded = GetOrCreateWorksheet(wbDash, SHEET_ADDED_TXNS)
    Set wsDropped = GetOrCreateWorksheet(wbDash, SHEET_DROPPED_TXNS)

    FormatAuditSheetHeaders wsAdded, "NEW TRANSACTIONS (IN NEW FILE, NOT IN OLD FILE)", AUDIT_ADDED
    FormatAuditSheetHeaders wsDropped, "DROPPED / EXCLUDED TRANSACTIONS (IN OLD FILE, NOT IN NEW FILE)", AUDIT_DROPPED

    stepName = "Clearing Old Dashboard Rows"
    ClearDashboardTable wsDash

    stepName = "Comparing Alert Pairs"
    curRow = FIRST_DATA_ROW
    For Each job In jobs
        jobNo = jobNo + 1
        Application.StatusBar = "AML comparison " & jobNo & " of " & jobs.Count & ": " & CStr(job(1))
        Select Case CStr(job(4))
            Case "PAIR"
                ProcessPair wsDash, wsAdded, wsDropped, curRow, CStr(job(0)), CStr(job(1)), _
                            CStr(job(2)), CStr(job(3)), narrativeDir, tot
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

    ReleaseWord

    stepName = "Formatting Dashboard Table"
    FormatBatchTable wsDash, FIRST_DATA_ROW, curRow - 1
    FinishAuditSheet wsAdded, 10, "No new transactions."
    FinishAuditSheet wsDropped, 9, "No dropped transactions."

    stepName = "Building New Counterparty Pivots"
    BuildNewCPPivotSheet wbDash, wsAdded, tot

    stepName = "Updating KPI Cards"
    UpdateKPICards wsDash, tot

    wsDash.Activate
    RestoreApp st

    summary = modeLabel & " Comparison Complete!" & vbCrLf & vbCrLf & _
              "- Alert pairs compared: " & tot.Evaluated & vbCrLf & _
              "- Unmatched files: " & tot.Unmatched & "   Duplicate files: " & tot.DupFiles & "   Errors: " & tot.Errors & vbCrLf & _
              "- Alerts with new transactions: " & tot.PairsWithNewTxns & vbCrLf & _
              "- New transactions: " & tot.Added & " (alerted " & tot.AddedAlerted & ")" & vbCrLf & _
              "- New transaction amount: " & Format(tot.AddedAmount, "#,##0.00") & " " & tot.Currencies & vbCrLf & _
              "- New counterparties identified: " & tot.NewCP & " (in " & tot.PairsWithNewCP & " alerts)" & vbCrLf & _
              "- Dropped transactions: " & tot.Dropped
    If narrativeDir <> "" Then
        summary = summary & vbCrLf & vbCrLf & "Narratives: updated " & tot.NarrUpdated & ", need review " & tot.NarrFlagged & _
                  ", no change " & tot.NarrNoChange & ", not found " & tot.NarrNotFound & ", errors " & tot.NarrErrors
        If tot.NarrOutDir <> "" Then summary = summary & vbCrLf & "Saved in: " & tot.NarrOutDir & vbCrLf & _
                  "Changed text is highlighted in yellow. See the 'Narrative Update' column for each alert."
    End If
    If InStr(tot.Currencies, ",") > 0 Then summary = summary & vbCrLf & vbCrLf & "WARNING: mixed currencies - amounts are not comparable."
    If tot.Errors > 0 Then summary = summary & vbCrLf & vbCrLf & "Some files could not be compared - see rows marked [ERROR]."
    MsgBox summary, IIf(tot.Errors > 0, vbExclamation, vbInformation), "Reconciliation Finished"
    Exit Sub

ErrorHandler:
    ReleaseWord
    RestoreApp st
    MsgBox "Error during comparison (" & stepName & "):" & vbCrLf & Err.Description, vbCritical, "Error " & Err.Number
End Sub
' =========================================================================
' [SEARCH] 7. RECONCILE ONE OLD/NEW WORKBOOK PAIR
' =========================================================================
Private Sub ProcessPair(ByVal wsDash As Worksheet, ByVal wsAdded As Worksheet, ByVal wsDropped As Worksheet, _
                        ByVal rowIdx As Long, ByVal ecmID As String, ByVal alertID As String, _
                        ByVal oldFilePath As String, ByVal newFilePath As String, _
                        ByVal narrativeDir As String, ByRef tot As BatchTotals)
    Dim oldSet As TxnSet, newSet As TxnSet
    Dim pr As PairResult
    Dim ctx As NarrCtx
    Dim oldCP As Object
    Dim statusStr As String, stage As String, narrStatus As String
    Dim rowVals(1 To 1, 1 To 18) As Variant

    On Error GoTo PairFail

    ' Each file is read into memory and closed before the next one is opened,
    ' so Old and New files may share the same file name.
    stage = "reading Old file"
    Set oldCP = CreateLookupDict()
    LoadTxnFile oldFilePath, oldSet, oldCP

    stage = "reading New file"
    LoadTxnFile newFilePath, newSet, Nothing

    stage = "comparing transactions"
    CompareSets wsAdded, wsDropped, ecmID, alertID, oldSet, newSet, oldCP, pr

    statusStr = BuildStatus(oldSet, newSet, pr)

    narrStatus = "-"
    If narrativeDir <> "" Then
        stage = "updating the narrative"
        Application.StatusBar = "Updating narrative for " & ecmID & "_" & alertID & "..."
        BuildNarrCtx ctx, ecmID, alertID, oldFilePath, newFilePath, oldSet, newSet, pr
        narrStatus = UpdateNarrative(ctx, narrativeDir, tot)
    End If

    stage = "writing dashboard row"

    rowVals(1, 1) = ecmID
    rowVals(1, 2) = alertID
    rowVals(1, 3) = GetFileName(oldFilePath) & "  [" & oldSet.SheetName & "]"
    rowVals(1, 4) = GetFileName(newFilePath) & "  [" & newSet.SheetName & "]"
    rowVals(1, 5) = FormatPeriod(oldSet.MinDate, oldSet.MaxDate, oldFilePath)
    rowVals(1, 6) = FormatPeriod(newSet.MinDate, newSet.MaxDate, newFilePath)
    rowVals(1, 7) = oldSet.UniqueCount
    rowVals(1, 8) = newSet.UniqueCount
    rowVals(1, 9) = pr.AddedCount
    rowVals(1, 10) = IIf(newSet.HasAmount, pr.AddedAmount, "n/a")
    rowVals(1, 11) = IIf(newSet.HasAlert, pr.AddedAlerted, "n/a")
    rowVals(1, 12) = pr.DroppedCount
    If pr.AddedCount > 0 And pr.CPUnknown = pr.AddedCount Then
        rowVals(1, 13) = "n/a"
    Else
        rowVals(1, 13) = pr.NewCPDistinct
    End If
    rowVals(1, 14) = IIf(pr.NewCPNames = "", "-", pr.NewCPNames)
    rowVals(1, 15) = IIf(oldSet.HasAmount, oldSet.TotalAmt, "n/a")
    rowVals(1, 16) = IIf(newSet.HasAmount, newSet.TotalAmt, "n/a")
    rowVals(1, 17) = statusStr
    rowVals(1, 18) = narrStatus

    wsDash.Cells(rowIdx, 2).Resize(1, 6).NumberFormat = "@"
    wsDash.Cells(rowIdx, 15).NumberFormat = "@"
    wsDash.Cells(rowIdx, 18).Resize(1, 2).NumberFormat = "@"
    wsDash.Cells(rowIdx, 2).Resize(1, DASH_COLS).Value = rowVals

    ' Totals
    tot.Evaluated = tot.Evaluated + 1
    If pr.AddedCount > 0 Then tot.PairsWithNewTxns = tot.PairsWithNewTxns + 1
    If pr.NewCPDistinct > 0 Then tot.PairsWithNewCP = tot.PairsWithNewCP + 1
    tot.Added = tot.Added + pr.AddedCount
    tot.AddedAlerted = tot.AddedAlerted + pr.AddedAlerted
    tot.AddedAmount = tot.AddedAmount + pr.AddedAmount
    tot.Dropped = tot.Dropped + pr.DroppedCount
    tot.NewCP = tot.NewCP + pr.NewCPDistinct
    tot.Currencies = MergeList(tot.Currencies, oldSet.Currencies)
    tot.Currencies = MergeList(tot.Currencies, newSet.Currencies)
    tot.AlertRows.Add Array(ecmID, alertID, pr.AddedCount, pr.AddedAmount, pr.NewCPDistinct, pr.NewCPNames, pr.NewCPAgg)
    Exit Sub

PairFail:
    tot.Errors = tot.Errors + 1
    WriteStatusRow wsDash, rowIdx, ecmID, alertID, GetFileName(oldFilePath), GetFileName(newFilePath), _
                   "ERROR while " & stage & ": " & Err.Description & " [ERROR]"
End Sub
Private Function BuildStatus(ByRef oldSet As TxnSet, ByRef newSet As TxnSet, ByRef pr As PairResult) As String
    Dim parts As String, detail As String, sev As Long, ccyAll As String
    ' sev: 0 = OK, 1 = WARN, 2 = ALERT

    If pr.AddedCount > 0 Then
        If newSet.HasAmount Then detail = "amount " & Format(pr.AddedAmount, "#,##0.00")
        If pr.AddedAlerted > 0 Then
            If detail <> "" Then detail = detail & "; "
            detail = detail & pr.AddedAlerted & " alerted"
        End If
        If detail <> "" Then
            AddPart parts, "NEW TXNS " & pr.AddedCount & " (" & detail & ")"
        Else
            AddPart parts, "NEW TXNS " & pr.AddedCount
        End If
        If sev < 1 Then sev = 1

        If pr.NewCPDistinct > 0 Then
            AddPart parts, "NEW CP " & pr.NewCPDistinct
            sev = 2
        ElseIf pr.CPUnknown < pr.AddedCount Then
            AddPart parts, "NO NEW CP"
        End If
        If pr.CPUnknown > 0 Then
            AddPart parts, "CP CHECK N/A FOR " & pr.CPUnknown & " TXNS"
            If sev < 1 Then sev = 1
        End If
    Else
        AddPart parts, "NO NEW TXNS"
    End If

    If pr.DroppedCount > 0 Then
        AddPart parts, "DROPPED " & pr.DroppedCount
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
    Resume LoadCleanup

LoadCleanup:
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
    Dim colRu As Variant, ruleTxt As String, ruleDict As Object
    Dim txDisp As String, txK As String, alertStr As String
    Dim dt As Variant, amt As Variant, acc As String, descTxt As String
    Dim cp As String, drcr As String, ccy As String, origName As String, benName As String
    Dim d As Date, a As Double, ok As Boolean
    Dim ccyDict As Object

    Set ts.Map = CreateLookupDict()
    Set ccyDict = CreateLookupDict()
    Set ruleDict = CreateLookupDict()
    ts.SheetName = ws.Name
    ts.MinDate = Empty
    ts.MaxDate = Empty
    ts.AlertMin = Empty
    ts.AlertMax = Empty

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
    If cm.Rule > 0 Then colRu = ReadBlock(ws, hdrRow + 1, cm.Rule, lastRow, cm.Rule)

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

                ' Credit / debit split
                Select Case DeriveCounterparty(drcr, "C", "D")
                    Case "C"
                        ts.CrCount = ts.CrCount + 1
                        If Not IsEmpty(amt) Then ts.CrAmt = ts.CrAmt + Abs(CDbl(amt))
                    Case "D"
                        ts.DrCount = ts.DrCount + 1
                        If Not IsEmpty(amt) Then ts.DrAmt = ts.DrAmt + Abs(CDbl(amt))
                End Select

                ' Alerted transactions: amount, date range and rule name
                If alertStr = "Yes" Then
                    If Not IsEmpty(amt) Then ts.AlertAmt = ts.AlertAmt + Abs(CDbl(amt))
                    If VarType(dt) = vbDate Then
                        If IsEmpty(ts.AlertMin) Then
                            ts.AlertMin = dt
                            ts.AlertMax = dt
                        Else
                            If dt < ts.AlertMin Then ts.AlertMin = dt
                            If dt > ts.AlertMax Then ts.AlertMax = dt
                        End If
                    End If
                    If cm.Rule > 0 Then
                        ruleTxt = CellText(colRu(i, 1))
                        If ruleTxt <> "" Then
                            If Not DictExists(ruleDict, ruleTxt) Then DictAdd ruleDict, ruleTxt, ruleTxt
                        End If
                    End If
                End If

                DictAdd ts.Map, txK, Array(txDisp, alertStr, dt, amt, acc, descTxt, cp, drcr, ccy)
            End If
        End If
    Next i

    ts.UniqueCount = DictCount(ts.Map)
    ts.Currencies = JoinSortedKeys(ccyDict)
    If DictCount(ruleDict) > 0 Then ts.RuleText = Join(DictKeys(ruleDict), "; ")
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
Private Sub CompareSets(ByVal wsAdded As Worksheet, ByVal wsDropped As Worksheet, _
                        ByVal ecmID As String, ByVal alertID As String, _
                        ByRef oldSet As TxnSet, ByRef newSet As TxnSet, ByVal oldCP As Object, _
                        ByRef pr As PairResult)
    Dim txKeys As Variant, i As Long, k As String
    Dim rec As Variant, kv As Variant, agg As Variant
    Dim addedKeys As Collection, droppedKeys As Collection
    Dim outArr() As Variant
    Dim n As Long, r As Long, startRow As Long
    Dim cpFlag As String, cpKey As String, cpName As String
    Dim amt As Double
    Dim isNew As Boolean, oldHasCPData As Boolean
    Dim cpKeys As Variant, j As Long

    Set addedKeys = New Collection
    Set droppedKeys = New Collection
    Set pr.NewCPAgg = CreateLookupDict()
    oldHasCPData = oldSet.HasCP Or (DictCount(oldCP) > 0)

    txKeys = DictKeys(newSet.Map)
    For i = LBound(txKeys) To UBound(txKeys)
        k = CStr(txKeys(i))
        If Not DictExists(oldSet.Map, k) Then addedKeys.Add k
    Next i

    txKeys = DictKeys(oldSet.Map)
    For i = LBound(txKeys) To UBound(txKeys)
        k = CStr(txKeys(i))
        If Not DictExists(newSet.Map, k) Then droppedKeys.Add k
    Next i

    ' --- New transactions ---
    n = addedKeys.Count
    pr.AddedCount = n
    If n > 0 Then
        ReDim outArr(1 To n, 1 To 10)
        r = 0
        For Each kv In addedKeys
            r = r + 1
            rec = DictGet(newSet.Map, CStr(kv))
            If rec(F_ALERT) = "Yes" Then pr.AddedAlerted = pr.AddedAlerted + 1
            If IsEmpty(rec(F_AMT)) Then amt = 0 Else amt = Abs(CDbl(rec(F_AMT)))
            pr.AddedAmount = pr.AddedAmount + amt

            cpName = Trim$(CStr(rec(F_CP)))
            cpFlag = CPStatus(cpName, oldHasCPData, newSet.HasCP, oldCP, isNew)
            If isNew Then
                pr.NewCPTxns = pr.NewCPTxns + 1
                cpKey = NormCP(cpName)
                If DictExists(pr.NewCPAgg, cpKey) Then
                    agg = DictGet(pr.NewCPAgg, cpKey)
                    agg(1) = agg(1) + 1
                    agg(2) = agg(2) + amt
                    DictSet pr.NewCPAgg, cpKey, agg
                Else
                    DictAdd pr.NewCPAgg, cpKey, Array(cpName, 1, amt)
                    pr.NewCPDistinct = pr.NewCPDistinct + 1
                    If pr.NewCPNames = "" Then
                        pr.NewCPNames = cpName
                    Else
                        pr.NewCPNames = pr.NewCPNames & "; " & cpName
                    End If
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

            If r <= 25 Then
                If pr.AddedDetail <> "" Then pr.AddedDetail = pr.AddedDetail & vbCr
                pr.AddedDetail = pr.AddedDetail & "  " & DateText(rec(F_DATE)) & "  $" & UsAmount(amt) & "  " & cpName
                If rec(F_ALERT) = "Yes" Then pr.AddedDetail = pr.AddedDetail & "  [alerted]"
            ElseIf r = 26 Then
                pr.AddedDetail = pr.AddedDetail & vbCr & "  ... and " & (n - 25) & " more - see sheet " & SHEET_ADDED_TXNS
            End If
        Next kv

        cpKeys = DictKeys(pr.NewCPAgg)
        For j = LBound(cpKeys) To UBound(cpKeys)
            agg = DictGet(pr.NewCPAgg, CStr(cpKeys(j)))
            If pr.NewCPDetail <> "" Then pr.NewCPDetail = pr.NewCPDetail & "; "
            pr.NewCPDetail = pr.NewCPDetail & agg(0) & " (" & agg(1) & " txn(s), $" & UsAmount(CDbl(agg(2))) & ")"
        Next j

        startRow = WriteAuditBlock(wsAdded, outArr, n, AUDIT_ADDED)
        For r = 1 To n
            If outArr(r, 4) = "Yes" Then HighlightCell wsAdded.Cells(startRow + r - 1, 5), COLOR_ALERT_RED, COLOR_FILL_RED
            If outArr(r, 10) = CP_YES Then
                wsAdded.Cells(startRow + r - 1, 10).Font.Bold = True
                HighlightCell wsAdded.Cells(startRow + r - 1, 11), COLOR_ALERT_RED, COLOR_FILL_RED
            ElseIf Left$(CStr(outArr(r, 10)), 7) = "Unknown" Then
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
End Sub
Private Function CPStatus(ByVal cp As String, ByVal oldHasCPData As Boolean, ByVal newHasCP As Boolean, _
                          ByVal oldCP As Object, ByRef isNew As Boolean) As String
    Dim cpKey As String
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
        CPStatus = CP_NO
    Else
        isNew = True
        CPStatus = CP_YES
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
    ws.Columns("J").ColumnWidth = 11 ' New Txns
    ws.Columns("K").ColumnWidth = 16 ' New Txn Amount
    ws.Columns("L").ColumnWidth = 11 ' New Txns Alerted
    ws.Columns("M").ColumnWidth = 11 ' Dropped Txns
    ws.Columns("N").ColumnWidth = 11 ' New CPs
    ws.Columns("O").ColumnWidth = 50 ' New CP Names
    ws.Columns("P").ColumnWidth = 16 ' Old Volume
    ws.Columns("Q").ColumnWidth = 16 ' New Volume
    ws.Columns("R").ColumnWidth = 60 ' Status Badge
    ws.Columns("S").ColumnWidth = 70 ' Narrative Update

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
    ws.Range("B4:S4").Merge
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

    ws.Range("B5:S5").Merge
    With ws.Range("B5")
        .Value = "Lookback period change check | New transactions (count and amount) and newly identified counterparties per alert"
        .Font.Name = FONT_UI
        .Font.Size = 9.5
        .Font.Color = RGB(100, 116, 139) ' Slate 500
        .Interior.Color = RGB(241, 245, 249) ' Slate 100
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' KPI Tiles (Columns B to R)
    CreateKPITile ws, "B7:D9", "ALERTS EVALUATED", "Alert pairs compared"
    CreateKPITile ws, "E7:G9", "ALERTS WITH NEW TXNS", "New file has extra transactions"
    CreateKPITile ws, "H7:J9", "NEW TXNS", "Present in New file only"
    CreateKPITile ws, "K7:M9", "NEW TXN AMOUNT", "Sum of new transaction amounts"
    CreateKPITile ws, "N7:S9", "NEW COUNTERPARTIES", "Not seen anywhere in the Old file"
    ws.Range("K8").NumberFormat = "#,##0.00"

    ' Action Buttons (Row 12)
    CreateActionButton ws, 12, "B", "C", "[RUN] Run Batch Comparison", "Run_Batch_Comparison", COLOR_ACCENT_BLUE
    CreateActionButton ws, 12, "D", "D", "[OPEN] Compare Single Pair", "Run_Single_Pair_Comparison", RGB(79, 70, 229)
    CreateActionButton ws, 12, "E", "E", "[EXPORT] Export to Workbook", "Export_To_New_Workbook", RGB(16, 149, 193)
    CreateActionButton ws, 12, "F", "F", "[VIEW] New Txns", "View_Added_Transactions", RGB(13, 148, 136)
    CreateActionButton ws, 12, "G", "G", "[VIEW] New CP Pivots", "View_New_CP_Pivots", RGB(13, 148, 136)
    CreateActionButton ws, 12, "H", "K", "[VIEW] Dropped Txns", "View_Dropped_Transactions", RGB(13, 148, 136)
    CreateActionButton ws, 12, "L", "O", "[RESET] Reset Dashboard", "Clear_Comparison_Dashboard", RGB(100, 116, 139)

    ' Table Header Row
    headers = Array("ECM ID", "Alert ID", "Old File Name [sheet]", "New File Name [sheet]", _
                    "Old Txn Date Range (data | file name)", "New Txn Date Range (data | file name)", _
                    "Old Count", "New Count", "New Txns", "New Txn Amount", "New Txns Alerted", "Dropped Txns", _
                    "New CPs", "New Counterparty Names", "Old Volume", "New Volume", "Reconciliation Status", _
                    "Narrative Update")

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
    SetKPI ws, "B8", tot.Evaluated, IIf(tot.Errors > 0, COLOR_ALERT_RED, COLOR_TEXT_DARK)
    ws.Range("B9").Value = "Compared: " & tot.Evaluated & " | Unmatched: " & tot.Unmatched & _
                           " | Dup files: " & tot.DupFiles & " | Errors: " & tot.Errors

    SetKPI ws, "E8", tot.PairsWithNewTxns, IIf(tot.PairsWithNewTxns > 0, COLOR_AMBER, COLOR_TEXT_DARK)
    ws.Range("E9").Value = "Alerts with new counterparties: " & tot.PairsWithNewCP

    SetKPI ws, "H8", tot.Added, IIf(tot.Added > 0, COLOR_ACCENT_BLUE, COLOR_TEXT_DARK)
    ws.Range("H9").Value = "Alerted: " & tot.AddedAlerted & " | Dropped: " & tot.Dropped

    SetKPI ws, "K8", tot.AddedAmount, IIf(tot.AddedAmount > 0, COLOR_ACCENT_BLUE, COLOR_TEXT_DARK)
    ws.Range("K8").NumberFormat = "#,##0.00"
    If InStr(tot.Currencies, ",") > 0 Then
        ws.Range("K9").Value = "MIXED CURRENCIES (" & tot.Currencies & ") - not comparable"
        ws.Range("K9").Font.Color = COLOR_ALERT_RED
    Else
        ws.Range("K9").Value = "Sum of new transaction amounts" & IIf(tot.Currencies <> "", " (" & tot.Currencies & ")", "")
        ws.Range("K9").Font.Color = RGB(148, 163, 184)
    End If

    SetKPI ws, "N8", tot.NewCP, IIf(tot.NewCP > 0, COLOR_ALERT_RED, COLOR_TEXT_DARK)
    If tot.NewCP > 0 Then
        ws.Range("N9").Value = "In " & tot.PairsWithNewCP & " alert(s) - see sheet " & SHEET_NEWCP_PIVOTS
    Else
        ws.Range("N9").Value = "Not seen anywhere in the Old file"
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
    With ws.Range(ws.Cells(startRow, 8), ws.Cells(endRow, 10))                            ' Old, New, New Txns
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0"
    End With
    With ws.Range(ws.Cells(startRow, 11), ws.Cells(endRow, 11))                           ' New Txn Amount
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0.00"
    End With
    With ws.Range(ws.Cells(startRow, 12), ws.Cells(endRow, 14))                           ' Alerted, Dropped, New CPs
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0"
    End With
    ws.Range(ws.Cells(startRow, 15), ws.Cells(endRow, 15)).HorizontalAlignment = xlLeft   ' New CP names
    With ws.Range(ws.Cells(startRow, 16), ws.Cells(endRow, 17))                           ' Volumes
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0.00"
    End With
    With ws.Range(ws.Cells(startRow, 18), ws.Cells(endRow, 18))                           ' Status
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
    End With
    ws.Range(ws.Cells(startRow, 19), ws.Cells(endRow, 19)).HorizontalAlignment = xlLeft   ' Narrative update

    For r = startRow To endRow
        If CellNumber(ws.Cells(r, 10)) > 0 Then HighlightCell ws.Cells(r, 10), COLOR_ALERT_RED, COLOR_FILL_RED
        If CellNumber(ws.Cells(r, 11)) > 0 Then
            ws.Cells(r, 11).Font.Bold = True
            ws.Cells(r, 11).Font.Color = COLOR_ALERT_RED
        End If
        If CellNumber(ws.Cells(r, 12)) > 0 Then
            ws.Cells(r, 12).Font.Bold = True
            ws.Cells(r, 12).Font.Color = COLOR_ALERT_RED
        End If
        If CellNumber(ws.Cells(r, 13)) > 0 Then
            ws.Cells(r, 13).Font.Bold = True
            ws.Cells(r, 13).Font.Color = COLOR_AMBER
        End If
        If CellNumber(ws.Cells(r, 14)) > 0 Then
            HighlightCell ws.Cells(r, 14), COLOR_ALERT_RED, COLOR_FILL_RED
            ws.Cells(r, 15).Font.Bold = True
            ws.Cells(r, 15).Font.Color = COLOR_ALERT_RED
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

        ' Narrative update result
        statusTxt = CStr(ws.Cells(r, 19).Value)
        If Left$(statusTxt, 7) = "Updated" Then
            ws.Cells(r, 19).Font.Color = COLOR_SUCCESS_GREEN
        ElseIf Left$(statusTxt, 5) = "Error" Then
            ws.Cells(r, 19).Font.Color = COLOR_ALERT_RED
        ElseIf Left$(statusTxt, 6) = "Review" Or Left$(statusTxt, 9) = "Narrative" Then
            ws.Cells(r, 19).Font.Color = COLOR_AMBER
        Else
            ws.Cells(r, 19).Font.Color = COLOR_TEXT_MUTED
        End If
    Next r
End Sub

' Numeric value of a cell, or 0 for blanks, text and errors
Private Function CellNumber(ByVal c As Range) As Double
    Dim v As Variant
    v = c.Value2
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger, vbByte
            CellNumber = CDbl(v)
    End Select
End Function
Private Sub HighlightCell(ByVal c As Range, ByVal fontColor As Long, ByVal fillColor As Long)
    c.Font.Bold = True
    c.Font.Color = fontColor
    c.Interior.Color = fillColor
End Sub

Private Sub WriteStatusRow(ByVal wsDash As Worksheet, ByVal rowIdx As Long, ByVal ecmID As String, ByVal alertID As String, _
                           ByVal oldF As String, ByVal newF As String, ByVal statusStr As String)
    Dim rowVals(1 To 1, 1 To 18) As Variant
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
    rowVals(1, 18) = "-"

    wsDash.Cells(rowIdx, 2).Resize(1, 6).NumberFormat = "@"
    wsDash.Cells(rowIdx, 18).Resize(1, 2).NumberFormat = "@"
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
    ws.Columns("E").ColumnWidth = 12 ' Alerted?
    ws.Columns("F").ColumnWidth = 14 ' Date
    ws.Columns("G").ColumnWidth = 16 ' Amount
    ws.Columns("H").ColumnWidth = 22 ' Account No
    ws.Columns("I").ColumnWidth = 28 ' Description
    ws.Columns("J").ColumnWidth = 34 ' Counterparty
    If kind = AUDIT_ADDED Then
        ws.Columns("K").ColumnWidth = 18 ' New Counterparty?
        lastColLetter = "K"
        headers = Array("ECM ID", "Alert ID", "Transaction ID", "Is Alerted?", "Transaction Date", "Amount", _
                        "Account No", "Transaction Description", "Counterparty Name", "New Counterparty?")
    Else
        lastColLetter = "J"
        headers = Array("ECM ID", "Alert ID", "Transaction ID", "Is Alerted?", "Transaction Date", "Amount", _
                        "Account No", "Transaction Description", "Counterparty Name")
    End If

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
Private Function WriteAuditBlock(ByVal ws As Worksheet, ByRef outArr() As Variant, ByVal nRows As Long, ByVal kind As Long) As Long
    Dim startRow As Long, nCols As Long
    Dim rng As Range

    startRow = LastUsedRow(ws) + 1
    If startRow < 4 Then startRow = 4
    nCols = UBound(outArr, 2)
    Set rng = ws.Cells(startRow, 2).Resize(nRows, nCols)

    ' Text columns first so IDs / account numbers keep leading zeros and are not turned into numbers
    ws.Cells(startRow, 2).Resize(nRows, 4).NumberFormat = "@"          ' ECM, Alert, Txn ID, Alerted?
    ws.Cells(startRow, 8).Resize(nRows, nCols - 6).NumberFormat = "@"  ' Account .. end

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
    ws.Cells(startRow, 4).Resize(nRows, 3).HorizontalAlignment = xlCenter   ' Txn ID, Alerted?, Date
    ws.Cells(startRow, 6).Resize(nRows, 1).NumberFormat = "yyyy-mm-dd"
    With ws.Cells(startRow, 7).Resize(nRows, 1)
        .NumberFormat = "#,##0.00;-#,##0.00;0.00"
        .HorizontalAlignment = xlRight
    End With
    If kind = AUDIT_ADDED Then ws.Cells(startRow, 11).Resize(nRows, 1).HorizontalAlignment = xlCenter

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
' [PIVOT] NEW COUNTERPARTY SUMMARY AND PIVOTS (sheet New_CP_Pivots)
' =========================================================================
' Top: one summary row per compared alert.
' Below: one pivot table per alert with new counterparties, built on the
' Discrepancy_Added_Txns sheet and filtered to that alert and
' "New Counterparty?" = Yes. Rows = counterparty, values = count and amount.
' If Excel cannot create a pivot (e.g. an older Mac build), a plain table
' with the same content is written instead.
Private Sub BuildNewCPPivotSheet(ByVal wb As Workbook, ByVal wsAdded As Worksheet, ByRef tot As BatchTotals)
    Dim ws As Worksheet
    Dim item As Variant
    Dim r As Long, nextRow As Long, lastAdded As Long, blockNo As Long, bottomRow As Long
    Dim srcAddr As String

    Set ws = GetOrCreateWorksheet(wb, SHEET_NEWCP_PIVOTS)
    ClearPivotSheet ws

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 70

    ws.Rows(1).RowHeight = 28
    ws.Range("B1:G1").Merge
    With ws.Range("B1")
        .Value = "NEWLY IDENTIFIED COUNTERPARTIES PER ALERT"
        .Font.Name = FONT_UI
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' --- Summary table ---
    WriteHeaderRow ws, 3, 2, Array("ECM ID", "Alert ID", "New Txns", "New Txn Amount", "New CPs", "New Counterparty Names")
    r = 4
    For Each item In tot.AlertRows
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 3).NumberFormat = "@"
        ws.Cells(r, 2).Value = item(0)
        ws.Cells(r, 3).Value = item(1)
        ws.Cells(r, 4).Value = item(2)
        ws.Cells(r, 5).Value = item(3)
        ws.Cells(r, 6).Value = item(4)
        ws.Cells(r, 7).Value = IIf(CStr(item(5)) = "", "-", item(5))
        If item(4) > 0 Then
            ws.Cells(r, 6).Font.Bold = True
            ws.Cells(r, 6).Font.Color = COLOR_ALERT_RED
            ws.Cells(r, 7).Font.Color = COLOR_ALERT_RED
        End If
        r = r + 1
    Next item
    If r = 4 Then
        ws.Cells(4, 2).Value = "No alert pairs were compared."
        r = 5
    Else
        With ws.Range(ws.Cells(4, 2), ws.Cells(r - 1, 7))
            .Font.Name = FONT_UI
            .Font.Size = 9
            With .Borders
                .LineStyle = xlContinuous
                .Color = COLOR_BORDER
                .Weight = xlThin
            End With
        End With
        ws.Range(ws.Cells(4, 4), ws.Cells(r - 1, 4)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(4, 5), ws.Cells(r - 1, 5)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(4, 6), ws.Cells(r - 1, 6)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(4, 2), ws.Cells(r - 1, 2)).HorizontalAlignment = xlCenter
    End If

    ' --- One pivot per alert with new counterparties ---
    lastAdded = LastUsedRow(wsAdded)
    If lastAdded >= 4 Then
        srcAddr = "'" & wsAdded.Name & "'!" & _
                  wsAdded.Range(wsAdded.Cells(3, 2), wsAdded.Cells(lastAdded, 11)).Address(ReferenceStyle:=xlR1C1)
    End If

    nextRow = r + 2
    For Each item In tot.AlertRows
        If item(4) > 0 Then
            blockNo = blockNo + 1
            With ws.Cells(nextRow, 2)
                .Value = "ECM " & item(0) & "  |  " & item(1) & "  -  " & item(4) & " new counterpart" & IIf(item(4) = 1, "y", "ies")
                .Font.Name = FONT_UI
                .Font.Size = 10
                .Font.Bold = True
                .Font.Color = COLOR_ALERT_RED
            End With

            ' Pivot body starts 6 rows down, leaving room above it for the 3 filter fields
            bottomRow = 0
            If srcAddr <> "" Then
                bottomRow = TryCreateNewCPPivot(wb, ws, srcAddr, nextRow + 6, "NewCP_" & blockNo, CStr(item(0)), CStr(item(1)))
            End If
            If bottomRow = 0 Then bottomRow = WriteStaticCPBlock(ws, nextRow + 2, item(6))
            nextRow = bottomRow + 3
        End If
    Next item

    If blockNo = 0 Then
        With ws.Cells(nextRow, 2)
            .Value = "No new counterparties identified."
            .Font.Name = FONT_UI
            .Font.Italic = True
            .Font.Color = RGB(148, 163, 184)
        End With
    End If
End Sub

' Creates one filtered pivot. Returns its last row, or 0 if Excel could not build it.
Private Function TryCreateNewCPPivot(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal srcAddr As String, _
                                     ByVal destRow As Long, ByVal ptName As String, _
                                     ByVal ecmID As String, ByVal alertID As String) As Long
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim df As PivotField

    On Error GoTo Fail
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcAddr)
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Cells(destRow, 2), TableName:=ptName)

    pt.ManualUpdate = True
    With pt.PivotFields("ECM ID")
        .Orientation = xlPageField
        .Position = 1
    End With
    With pt.PivotFields("Alert ID")
        .Orientation = xlPageField
        .Position = 2
    End With
    With pt.PivotFields("New Counterparty?")
        .Orientation = xlPageField
        .Position = 3
    End With
    With pt.PivotFields("Counterparty Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    pt.AddDataField pt.PivotFields("Transaction ID"), "No of Trx", xlCount
    Set df = pt.AddDataField(pt.PivotFields("Amount"), "Sum of Transaction Amount", xlSum)
    df.NumberFormat = "#,##0.00"

    pt.PivotFields("ECM ID").CurrentPage = ecmID
    pt.PivotFields("Alert ID").CurrentPage = alertID
    pt.PivotFields("New Counterparty?").CurrentPage = CP_YES
    pt.ManualUpdate = False

    ' Cosmetics only - ignore if this Excel build does not support them
    On Error Resume Next
    pt.CompactLayoutRowHeader = "New Counterparty"
    pt.TableStyle2 = "PivotStyleLight16"
    pt.PivotFields("Counterparty Name").AutoSort xlDescending, "Sum of Transaction Amount"
    On Error GoTo Fail

    With pt.TableRange2
        TryCreateNewCPPivot = .Row + .Rows.Count - 1
    End With
    Exit Function

Fail:
    Resume FailCleanup

FailCleanup:
    On Error Resume Next
    If Not pt Is Nothing Then pt.TableRange2.Clear
    TryCreateNewCPPivot = 0
End Function

' Plain-table fallback with the same content as the pivot. Returns its last row.
Private Function WriteStaticCPBlock(ByVal ws As Worksheet, ByVal startRow As Long, ByVal agg As Object) As Long
    Dim keys As Variant, v As Variant
    Dim i As Long, r As Long, totCount As Long, totAmt As Double

    WriteHeaderRow ws, startRow, 2, Array("New Counterparty", "No of Trx", "Sum of Transaction Amount")
    r = startRow + 1
    If Not agg Is Nothing Then
        keys = DictKeys(agg)
        For i = LBound(keys) To UBound(keys)
            v = DictGet(agg, CStr(keys(i)))
            ws.Cells(r, 2).NumberFormat = "@"
            ws.Cells(r, 2).Value = v(0)
            ws.Cells(r, 3).Value = v(1)
            ws.Cells(r, 4).Value = v(2)
            totCount = totCount + v(1)
            totAmt = totAmt + v(2)
            r = r + 1
        Next i
    End If
    ws.Cells(r, 2).Value = "Grand Total"
    ws.Cells(r, 3).Value = totCount
    ws.Cells(r, 4).Value = totAmt
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Font.Bold = True

    With ws.Range(ws.Cells(startRow + 1, 2), ws.Cells(r, 4))
        .Font.Name = FONT_UI
        .Font.Size = 9
        With .Borders
            .LineStyle = xlContinuous
            .Color = COLOR_BORDER
            .Weight = xlThin
        End With
    End With
    ws.Range(ws.Cells(startRow + 1, 3), ws.Cells(r, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(startRow + 1, 4), ws.Cells(r, 4)).NumberFormat = "#,##0.00"
    WriteStaticCPBlock = r
End Function

Private Sub WriteHeaderRow(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal firstCol As Long, ByVal headers As Variant)
    Dim c As Long
    For c = 0 To UBound(headers)
        With ws.Cells(rowIdx, firstCol + c)
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

Private Sub ClearPivotSheet(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.PivotTables.Count To 1 Step -1
        ws.PivotTables(i).TableRange2.Clear
    Next i
    On Error GoTo 0
    ws.Cells.UnMerge
    ws.Cells.Clear
End Sub

' =========================================================================
' [NARRATIVE] UPDATE ESCALATION NARRATIVES WITH NEW FILE FIGURES (Word)
' =========================================================================
' For each compared alert, the narrative ECMID_ALERTID_*.docx is copied into
' a new "Updated_Narratives_<timestamp>" subfolder of the narrative folder and
' the copy is edited in Word. The original is never changed.
'
' Only these five lines are updated, with values taken from the New file:
'  1. Opening sentence ("... to report approximately N credit transactions
'     totaling $X ...")                 -> alerted count and alerted total
'  2. "Total Suspicious Dollar Amount"   -> alerted total
'  3. "Date Range of Suspicious Activity"-> first and last alerted date
'  4. First sentence of "Between D1 and D2, ... N transactions totaling $X"
'  5. First sentence of "Specifically, between D1 through D2, ... N credit
'     transactions totaling $X to K counterparties ..."
'     Lines 4-5: each date, count and amount is replaced with the matching New
'     file figure when it equals the Old file's figure (overall, credits or
'     debits; first/last transaction date or file-name period). A figure that
'     matches neither file is left alone and commented. K is increased by the
'     number of newly identified counterparties.
' Lines 1-3 always take the New file's alerted figures.
' The new value is written in the existing font (not tracked) and highlighted
' yellow. Anything that cannot be located is reported as a Word comment.
Private Function UpdateNarrative(ByRef ctx As NarrCtx, ByVal narrativeDir As String, ByRef tot As BatchTotals) As String
    Dim srcPath As String, outPath As String, note As String, stage As String, errDesc As String
    Dim doc As Object
    Dim edits As Collection, flags As Collection
    Dim copied As Boolean, wasTracking As Boolean
    Dim status As String

    On Error GoTo Fail
    stage = "finding the narrative"
    srcPath = FindNarrativeFile(narrativeDir, ctx.EcmID, ctx.AlertID, note)
    If srcPath = "" Then
        tot.NarrNotFound = tot.NarrNotFound + 1
        UpdateNarrative = "Narrative not found (" & NarrativePrefix(ctx.EcmID, ctx.AlertID) & "_*.docx)"
        Exit Function
    End If
    If Not NarrativeNeedsUpdate(ctx) Then
        tot.NarrNoChange = tot.NarrNoChange + 1
        UpdateNarrative = "No update needed - Old and New file figures are the same"
        Exit Function
    End If

    stage = "copying the narrative"
    If tot.NarrOutDir = "" Then
        tot.NarrOutDir = narrativeDir & "Updated_Narratives_" & Format(Now, "yyyymmdd_hhnnss") & Application.PathSeparator
        MkDir Left$(tot.NarrOutDir, Len(tot.NarrOutDir) - 1)
    End If
    outPath = tot.NarrOutDir & GetFileName(srcPath)
    FileCopy srcPath, outPath
    copied = True

    stage = "opening the narrative in Word"
    Set doc = GetWord().Documents.Open(FileName:=outPath, ReadOnly:=False, AddToRecentFiles:=False)
    wasTracking = doc.TrackRevisions
    doc.TrackRevisions = False

    stage = "editing the narrative"
    Set edits = New Collection
    Set flags = New Collection
    PlanNarrativeEdits doc, ctx, edits, flags

    If edits.Count = 0 And flags.Count = 0 Then
        ' Narrative already shows the New file figures - no copy needed
        doc.Close SaveChanges:=0
        Set doc = Nothing
        Kill outPath
        tot.NarrNoChange = tot.NarrNoChange + 1
        UpdateNarrative = "No update needed - narrative already matches the New file"
        Exit Function
    End If

    ApplyNarrativeEdits doc, edits
    If NARRATIVE_ADD_COMMENTS Or flags.Count > 0 Then AddNarrativeComments doc, ctx, flags, edits.Count, note

    stage = "saving the narrative"
    doc.TrackRevisions = wasTracking
    doc.Save
    doc.Close SaveChanges:=0
    Set doc = Nothing

    If edits.Count > 0 Then
        tot.NarrUpdated = tot.NarrUpdated + 1
        status = "Updated: " & edits.Count & " value(s) changed, highlighted yellow"
    Else
        tot.NarrFlagged = tot.NarrFlagged + 1
        status = "Review: lines not found - see comments"
    End If
    If flags.Count > 0 And edits.Count > 0 Then status = status & "; " & flags.Count & " item(s) to review in comments"
    UpdateNarrative = status & " -> " & GetFileName(Left$(tot.NarrOutDir, Len(tot.NarrOutDir) - 1)) & _
                      Application.PathSeparator & GetFileName(outPath)
    Exit Function

Fail:
    errDesc = Err.Description
    Resume FailCleanup

FailCleanup:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=0
    If copied Then Kill outPath
    On Error GoTo 0
    tot.NarrErrors = tot.NarrErrors + 1
    UpdateNarrative = "Error while " & stage & ": " & errDesc
End Function

Private Function NarrativeNeedsUpdate(ByRef ctx As NarrCtx) As Boolean
    NarrativeNeedsUpdate = (ctx.OldCount <> ctx.NewCount) Or (Abs(ctx.OldTotal - ctx.NewTotal) > 0.005) _
        Or (UsDate(ctx.OldMin) <> UsDate(ctx.NewMin)) Or (UsDate(ctx.OldMax) <> UsDate(ctx.NewMax)) _
        Or (ctx.OldAlertCount <> ctx.NewAlertCount) Or (Abs(ctx.OldAlertAmt - ctx.NewAlertAmt) > 0.005) _
        Or (UsDate(ctx.OldAlertMin) <> UsDate(ctx.NewAlertMin)) Or (UsDate(ctx.OldAlertMax) <> UsDate(ctx.NewAlertMax)) _
        Or (ctx.AddedCount > 0) Or (ctx.DroppedCount > 0) Or (ctx.NewCPDistinct > 0)
End Function

Private Function NarrativePrefix(ByVal ecmID As String, ByVal alertID As String) As String
    If ecmID <> "-" And ecmID <> "" Then
        NarrativePrefix = ecmID & "_" & alertID
    Else
        NarrativePrefix = alertID
    End If
End Function

' Finds ECMID_ALERTID*.docx (top level of the folder only). With several
' matches the most recently modified one is used and noted.
Private Function FindNarrativeFile(ByVal folderPath As String, ByVal ecmID As String, ByVal alertID As String, _
                                   ByRef note As String) As String
    Dim prefix As String, sFile As String, ext As String, nextCh As String
    Dim matches As New Collection
    Dim v As Variant, bestPath As String, bestTime As Date, t As Date

    note = ""
    prefix = NarrativePrefix(ecmID, alertID)
    sFile = Dir(folderPath, vbReadOnly)
    Do While sFile <> ""
        ext = LCase$(Mid$(sFile, InStrRev(sFile, ".") + 1))
        If (ext = "docx" Or ext = "docm" Or ext = "doc") And Left$(sFile, 2) <> "~$" And Left$(sFile, 2) <> "._" Then
            If StrComp(Left$(sFile, Len(prefix)), prefix, vbTextCompare) = 0 Then
                nextCh = Mid$(sFile, Len(prefix) + 1, 1)
                If Not (nextCh Like "#") Then matches.Add folderPath & sFile
            End If
        End If
        sFile = Dir()
    Loop

    If matches.Count = 0 Then Exit Function
    For Each v In matches
        t = FileDateTime(CStr(v))
        If bestPath = "" Or t > bestTime Then
            bestPath = CStr(v)
            bestTime = t
        End If
    Next v
    If matches.Count > 1 Then note = matches.Count & " narratives match " & prefix & "; the most recently modified one was used."
    FindNarrativeFile = bestPath
End Function

' Locates the five lines and plans value replacements (Array(start, end, oldText, newText)).
Private Sub PlanNarrativeEdits(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal edits As Collection, ByVal flags As Collection)
    Dim p As Object
    Dim t As String, lt As String
    Dim escS As Long, totS As Long, drS As Long, betS As Long, spcS As Long
    Dim escT As String, totT As String, drT As String, betT As String, spcT As String
    Dim newCount As String, newTotal As String, newAlertCount As String, newAlertTotal As String
    Dim sentEnd As Long, k As Long, kLen As Long, kTxt As String

    escS = -1: totS = -1: drS = -1: betS = -1: spcS = -1
    For Each p In doc.Paragraphs
        t = p.Range.Text
        lt = LCase$(LTrim$(t))
        If escS < 0 And InStr(lt, "escalated") > 0 And InStr(lt, "totaling") > 0 Then
            escS = p.Range.Start
            escT = t
        ElseIf totS < 0 And InStr(lt, "total suspicious dollar amount") > 0 Then
            totS = p.Range.Start
            totT = t
        ElseIf drS < 0 And InStr(lt, "date range of suspicious activity") > 0 Then
            drS = p.Range.Start
            drT = t
        ElseIf betS < 0 And Left$(lt, 8) = "between " And InStr(lt, "totaling") > 0 Then
            betS = p.Range.Start
            betT = t
        ElseIf spcS < 0 And Left$(lt, 12) = "specifically" And InStr(lt, "between") > 0 And InStr(lt, "totaling") > 0 Then
            spcS = p.Range.Start
            spcT = t
        End If
    Next p

    newCount = UsNumber(ctx.NewCount)
    newTotal = "$" & UsAmount(ctx.NewTotal)
    newAlertCount = UsNumber(ctx.NewAlertCount)
    newAlertTotal = "$" & UsAmount(ctx.NewAlertAmt)

    ' Lines 1-3: alerted activity (needs the New file's "Is Alerted" column)
    If ctx.NewAlertCount = 0 Then
        flags.Add Array("", "The New file has no alerted transactions (or no 'Is Alerted' column), so the opening sentence, " & _
                  "Total Suspicious Dollar Amount and Date Range of Suspicious Activity were not changed.")
    Else
        If escS >= 0 Then
            sentEnd = SentenceEnd(escT, 1)
            PlanCount doc, edits, escS, escT, 1, sentEnd, Array("transaction", "transfer", "txn"), newAlertCount
            PlanAmount doc, edits, escS, escT, InStr(1, escT, "totaling", vbTextCompare), sentEnd, newAlertTotal
        Else
            flags.Add Array("", "Opening sentence ('... to report N transactions totaling $X') not found. New file alerted: " & _
                      newAlertCount & " transactions totaling " & newAlertTotal & ".")
        End If
        If totS >= 0 Then
            PlanAmount doc, edits, totS, totT, 1, Len(totT), newAlertTotal
        Else
            flags.Add Array("", "'Total Suspicious Dollar Amount' not found. New file alerted total: " & newAlertTotal & ".")
        End If
        If drS >= 0 Then
            PlanDates doc, edits, flags, drS, drT, 1, Len(drT), UsDate(ctx.NewAlertMin), UsDate(ctx.NewAlertMax), _
                      "Date Range of Suspicious Activity"
        Else
            flags.Add Array("", "'Date Range of Suspicious Activity' not found. New file alerted dates: " & _
                      UsDate(ctx.NewAlertMin) & " through " & UsDate(ctx.NewAlertMax) & ".")
        End If
    End If

    ' Lines 4 and 5 (the "Specifically" line is optional)
    If betS >= 0 Then PlanActivityLine doc, ctx, edits, flags, betS, betT, "'Between ...' line"
    If spcS >= 0 Then
        PlanActivityLine doc, ctx, edits, flags, spcS, spcT, "'Specifically ...' line"
        sentEnd = SentenceEnd(spcT, 1)

        ' Counterparty count: add the counterparties identified as new in the New file
        If ctx.NewCPDistinct > 0 Then
            k = FindCountToken(spcT, 1, sentEnd, Array("counterpart"), kLen)
            If k > 0 Then
                kTxt = Mid$(spcT, k, kLen)
                AddValueEdit doc, edits, spcS, spcT, k, kTxt, UsNumber(CLng(Replace(kTxt, ",", "")) + ctx.NewCPDistinct)
            Else
                flags.Add Array("Specifically", ctx.NewCPDistinct & " newly identified counterpart" & IIf(ctx.NewCPDistinct = 1, "y", "ies") & _
                          " - the counterparty count was not found in this line; please update it.")
            End If
        End If
        If ctx.DroppedCount > 0 Then
            flags.Add Array("Specifically", ctx.DroppedCount & " transaction(s) from the Old file are not in the New file; " & _
                      "check whether the counterparty count needs to go down.")
        End If
    End If
    If betS < 0 And spcS < 0 Then
        flags.Add Array("", "No 'Between [date] and [date], ... N transactions totaling $X' line was found. New file: " & _
                  newCount & " transactions totaling " & newTotal & " between " & UsDate(ctx.NewMin) & " and " & _
                  UsDate(ctx.NewMax) & ".")
    End If
End Sub

' First two dates of a line -> new first/last date
Private Sub PlanDates(ByVal doc As Object, ByVal edits As Collection, ByVal flags As Collection, ByVal parStart As Long, _
                      ByVal t As String, ByVal fromPos As Long, ByVal toPos As Long, ByVal newFrom As String, _
                      ByVal newTo As String, ByVal lineName As String)
    Dim p1 As Long, l1 As Long, p2 As Long, l2 As Long
    If newFrom = "" Or newTo = "" Then Exit Sub
    p1 = FindDateToken(t, fromPos, toPos, l1)
    If p1 > 0 Then p2 = FindDateToken(t, p1 + l1, toPos, l2)
    If p1 = 0 Or p2 = 0 Then
        flags.Add Array("", lineName & ": the two dates were not found. New file dates: " & newFrom & " through " & newTo & ".")
        Exit Sub
    End If
    AddValueEdit doc, edits, parStart, t, p1, Mid$(t, p1, l1), newFrom
    AddValueEdit doc, edits, parStart, t, p2, Mid$(t, p2, l2), newTo
End Sub

' Dates, then every "N [credit/debit] transactions totaling $X" pair in the first sentence of a line.
' All or nothing: if any figure matches neither file the line is left unchanged and commented.
Private Sub PlanActivityLine(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal edits As Collection, ByVal flags As Collection, _
                             ByVal parStart As Long, ByVal t As String, ByVal lineName As String)
    Dim sentEnd As Long, p1 As Long, l1 As Long, p2 As Long, l2 As Long
    Dim pos As Long, p As Long, n As Long, nextP As Long, nn As Long, a As Long, an As Long, bound As Long
    Dim hint As String, followWords As Variant, msg As String
    Dim lineEdits As Collection, lineFlags As Collection, v As Variant, w As Variant, overlaps As Boolean

    Set lineEdits = New Collection
    Set lineFlags = New Collection
    sentEnd = SentenceEnd(t, 1)
    followWords = Array("transaction", "transfer", "txn", "credit", "debit")

    p1 = FindDateToken(t, 1, sentEnd, l1)
    If p1 > 0 Then p2 = FindDateToken(t, p1 + l1, sentEnd, l2)
    If p1 > 0 And p2 > 0 Then
        PlanMapped doc, lineEdits, lineFlags, parStart, t, p1, l1, _
                   Array(Array(UsDate(ctx.OldMin), UsDate(ctx.NewMin)), Array(UsDate(ctx.OldPeriodFrom), UsDate(ctx.NewPeriodFrom))), _
                   lineName, "start date"
        PlanMapped doc, lineEdits, lineFlags, parStart, t, p2, l2, _
                   Array(Array(UsDate(ctx.OldMax), UsDate(ctx.NewMax)), Array(UsDate(ctx.OldPeriodTo), UsDate(ctx.NewPeriodTo))), _
                   lineName, "end date"
    Else
        lineFlags.Add Array("", "the date range was not found")
    End If

    pos = 1
    Do
        p = FindCountToken(t, pos, sentEnd, followWords, n)
        If p = 0 Then Exit Do
        hint = CountDirection(t, p, n)
        PlanMapped doc, lineEdits, lineFlags, parStart, t, p, n, CountPairs(ctx, hint), lineName, "transaction count"
        nextP = FindCountToken(t, p + n, sentEnd, followWords, nn)
        If nextP > 0 Then bound = nextP - 1 Else bound = sentEnd
        a = FindAmountToken(t, p + n, bound, an)
        If a > 0 Then PlanMapped doc, lineEdits, lineFlags, parStart, t, a, an, AmountPairs(ctx, hint), lineName, "amount"
        pos = p + n
    Loop

    If lineFlags.Count = 0 Then
        For Each v In lineEdits
            overlaps = False
            For Each w In edits
                If v(0) < w(1) And v(1) > w(0) Then overlaps = True
            Next w
            If Not overlaps Then edits.Add v
        Next v
    Else
        For Each v In lineFlags
            If msg <> "" Then msg = msg & "; "
            msg = msg & CStr(v(1))
        Next v
        msg = lineName & " was not changed because " & msg & ". New file: " & UsNumber(ctx.NewCount) & _
              " transactions totaling $" & UsAmount(ctx.NewTotal) & " between " & UsDate(ctx.NewMin) & " and " & UsDate(ctx.NewMax)
        If ctx.HasDrCr Then msg = msg & " (credits: " & UsNumber(ctx.NewCrCount) & " totaling $" & UsAmount(ctx.NewCrAmt) & _
                                  "; debits: " & UsNumber(ctx.NewDrCount) & " totaling $" & UsAmount(ctx.NewDrAmt) & ")"
        If IsDate(ctx.NewPeriodFrom) Then msg = msg & "; New file period " & UsDate(ctx.NewPeriodFrom) & " to " & UsDate(ctx.NewPeriodTo)
        flags.Add Array(Left$(Trim$(t), 40), msg & ".")
    End If
End Sub

' Replaces a figure with the New value of the first pair whose Old value it equals.
' Already-new figures are left alone; figures matching neither file are commented.
Private Sub PlanMapped(ByVal doc As Object, ByVal edits As Collection, ByVal flags As Collection, ByVal parStart As Long, _
                       ByVal t As String, ByVal pos As Long, ByVal n As Long, ByVal pairs As Variant, _
                       ByVal lineName As String, ByVal what As String)
    Dim cur As String, curKey As String, pr As Variant, oldList As String

    cur = Mid$(t, pos, n)
    curKey = NormValue(cur)
    For Each pr In pairs
        If CStr(pr(1)) <> "" Then
            If NormValue(CStr(pr(1))) = curKey Then Exit Sub
        End If
    Next pr
    For Each pr In pairs
        If CStr(pr(0)) <> "" And CStr(pr(1)) <> "" Then
            If NormValue(CStr(pr(0))) = curKey Then
                AddValueEdit doc, edits, parStart, t, pos, cur, CStr(pr(1))
                Exit Sub
            End If
            If oldList <> "" Then oldList = oldList & " / "
            oldList = oldList & CStr(pr(0))
        End If
    Next pr
    flags.Add Array("", what & " " & cur & " matches neither the Old file (" & oldList & ") nor the New file")
End Sub

' Old/New count pairs, most likely first: credits for "credit/incoming/received",
' debits for "debit/outgoing/sent", otherwise all transactions
Private Function CountPairs(ByRef ctx As NarrCtx, ByVal hint As String) As Variant
    Dim allP As Variant, crP As Variant, drP As Variant
    allP = Array(UsNumber(ctx.OldCount), UsNumber(ctx.NewCount))
    crP = Array(UsNumber(ctx.OldCrCount), UsNumber(ctx.NewCrCount))
    drP = Array(UsNumber(ctx.OldDrCount), UsNumber(ctx.NewDrCount))
    If Not ctx.HasDrCr Then
        CountPairs = Array(allP)
    ElseIf hint = "C" Then
        CountPairs = Array(crP, allP, drP)
    ElseIf hint = "D" Then
        CountPairs = Array(drP, allP, crP)
    Else
        CountPairs = Array(allP, crP, drP)
    End If
End Function

Private Function AmountPairs(ByRef ctx As NarrCtx, ByVal hint As String) As Variant
    Dim allP As Variant, crP As Variant, drP As Variant
    allP = Array("$" & UsAmount(ctx.OldTotal), "$" & UsAmount(ctx.NewTotal))
    crP = Array("$" & UsAmount(ctx.OldCrAmt), "$" & UsAmount(ctx.NewCrAmt))
    drP = Array("$" & UsAmount(ctx.OldDrAmt), "$" & UsAmount(ctx.NewDrAmt))
    If Not ctx.HasDrCr Then
        AmountPairs = Array(allP)
    ElseIf hint = "C" Then
        AmountPairs = Array(crP, allP, drP)
    ElseIf hint = "D" Then
        AmountPairs = Array(drP, allP, crP)
    Else
        AmountPairs = Array(allP, crP, drP)
    End If
End Function

' "C" when the count reads as credits ("376 credit ...", "received 26 ..."),
' "D" for debits ("10 debit ...", "sent 620 ..."), otherwise ""
Private Function CountDirection(ByVal t As String, ByVal pos As Long, ByVal n As Long) As String
    Dim rest As String, before As String, words() As String, i As Long, cnt As Long, w As String
    rest = LCase$(Mid$(t, pos + n, 40))
    words = Split(Application.WorksheetFunction.Trim(Replace(Replace(rest, ",", " "), ".", " ")), " ")
    For i = 0 To UBound(words)
        w = words(i)
        If w <> "" Then
            cnt = cnt + 1
            If Left$(w, 6) = "credit" Or Left$(w, 8) = "incoming" Or Left$(w, 7) = "inbound" Or Left$(w, 7) = "deposit" Then
                CountDirection = "C"
                Exit Function
            End If
            If Left$(w, 5) = "debit" Or Left$(w, 8) = "outgoing" Or Left$(w, 8) = "outbound" Or Left$(w, 10) = "withdrawal" Then
                CountDirection = "D"
                Exit Function
            End If
            If cnt >= 2 Then Exit For
        End If
    Next i
    If pos > 1 Then
        before = LCase$(Application.WorksheetFunction.Trim(Mid$(t, 1, pos - 1)))
        If Right$(before, 8) = "received" Or Right$(before, 9) = "receiving" Then
            CountDirection = "C"
        ElseIf Right$(before, 4) = "sent" Or Right$(before, 7) = "sending" Or Right$(before, 4) = "paid" Then
            CountDirection = "D"
        End If
    End If
End Function

' First number followed (within three words) by one of the given words -> new count
Private Sub PlanCount(ByVal doc As Object, ByVal edits As Collection, ByVal parStart As Long, ByVal t As String, _
                      ByVal fromPos As Long, ByVal toPos As Long, ByVal followWords As Variant, ByVal newCount As String)
    Dim p As Long, n As Long
    p = FindCountToken(t, fromPos, toPos, followWords, n)
    If p > 0 Then AddValueEdit doc, edits, parStart, t, p, Mid$(t, p, n), newCount
End Sub

' First "$amount" at or after fromPos -> new amount
Private Sub PlanAmount(ByVal doc As Object, ByVal edits As Collection, ByVal parStart As Long, ByVal t As String, _
                       ByVal fromPos As Long, ByVal toPos As Long, ByVal newAmount As String)
    Dim p As Long, n As Long
    If fromPos < 1 Then fromPos = 1
    p = FindAmountToken(t, fromPos, toPos, n)
    If p > 0 Then AddValueEdit doc, edits, parStart, t, p, Mid$(t, p, n), newAmount
End Sub

' Plans one replacement when the value differs. The document position is checked
' against the text; if they disagree the value is searched for inside the paragraph.
Private Sub AddValueEdit(ByVal doc As Object, ByVal edits As Collection, ByVal parStart As Long, ByVal t As String, _
                         ByVal offset As Long, ByVal oldText As String, ByVal newText As String)
    Dim s As Long, e As Long, hits As Collection, h As Variant, v As Variant

    If NormValue(oldText) = NormValue(newText) Then Exit Sub
    s = parStart + offset - 1
    e = s + Len(oldText)
    If doc.Range(s, e).Text <> oldText Then
        Set hits = WdFindIn(doc, parStart, parStart + Len(t) + 50, oldText, False)
        If hits.Count = 0 Then Exit Sub
        h = hits(1)
        s = h(0)
        e = h(1)
    End If
    For Each v In edits
        If s < v(1) And e > v(0) Then Exit Sub
    Next v
    edits.Add Array(s, e, oldText, newText)
End Sub

' Comparable form of a figure: dates as yyyymmdd ("6/2/2026" = "06/02/2026"),
' numbers as cents ("$60,990" = "$60,990.00"), anything else as typed
Private Function NormValue(ByVal s As String) As String
    Dim d As Date, ok As Boolean
    s = Replace(Replace(Replace(Replace(s, "$", ""), ",", ""), " ", ""), Chr(160), "")
    If s Like "*/*/*" Then
        d = ParseDateValue(s, ok)
        If ok Then
            NormValue = Format$(d, "yyyymmdd")
            Exit Function
        End If
    End If
    If s Like "*#*" And Not (s Like "*[!0-9.]*") Then
        NormValue = Format$(Val(s) * 100, "0")
    Else
        NormValue = s
    End If
End Function

' Applies edits from the end of the document backwards so earlier positions stay valid.
' The new value is inserted after the old one (so it takes the same font), highlighted
' yellow, and the old value is then removed. Nothing is tracked.
Private Sub ApplyNarrativeEdits(ByVal doc As Object, ByVal edits As Collection)
    Dim arr() As Variant, i As Long, j As Long, n As Long, tmp As Variant
    Dim ins As Object

    n = edits.Count
    If n = 0 Then Exit Sub
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = edits(i)
    Next i
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j)(0) >= tmp(0) Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i

    doc.TrackRevisions = False
    For i = 1 To n
        Set ins = doc.Range(arr(i)(1), arr(i)(1))
        ins.InsertAfter CStr(arr(i)(3))
        ins.HighlightColorIndex = 7          ' wdYellow
        doc.Range(arr(i)(0), arr(i)(1)).Delete
    Next i
End Sub

Private Sub AddNarrativeComments(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal flags As Collection, _
                                 ByVal nEdits As Long, ByVal note As String)
    Dim txt As String, docText As String, rulePart As Variant, f As Variant
    Dim ruleMissing As String

    If NARRATIVE_ADD_COMMENTS Then
        docText = doc.Content.Text

        ' Summary of what changed between the Old and New files
        txt = "LOOKBACK UPDATE (macro, " & Format(Now, "yyyy-mm-dd") & ")" & vbCr & _
              "Old file: " & ctx.OldFile & " - " & UsNumber(ctx.OldCount) & " transactions, $" & UsAmount(ctx.OldTotal) & _
              ", " & UsDate(ctx.OldMin) & " to " & UsDate(ctx.OldMax) & vbCr & _
              "New file: " & ctx.NewFile & " - " & UsNumber(ctx.NewCount) & " transactions, $" & UsAmount(ctx.NewTotal) & _
              ", " & UsDate(ctx.NewMin) & " to " & UsDate(ctx.NewMax) & vbCr & _
              "Alerted in New file: " & ctx.NewAlertCount & " transactions, $" & UsAmount(ctx.NewAlertAmt) & ", " & _
              UsDate(ctx.NewAlertMin) & " to " & UsDate(ctx.NewAlertMax) & vbCr & _
              "New transactions: " & ctx.AddedCount & " totaling $" & UsAmount(ctx.AddedAmount) & " (" & ctx.AddedAlerted & " alerted)"
        If ctx.AddedDetail <> "" Then txt = txt & vbCr & ctx.AddedDetail
        txt = txt & vbCr & "Dropped transactions: " & ctx.DroppedCount & vbCr & "Newly identified counterparties: "
        If ctx.NewCPDistinct > 0 Then
            txt = txt & ctx.NewCPDetail
        Else
            txt = txt & "none - every counterparty appears in the Old file"
        End If
        txt = txt & vbCr & nEdits & " value(s) changed; each is highlighted in yellow. Remove the highlight before filing."
        If note <> "" Then txt = txt & vbCr & note
        WdComment doc, "", txt

        ' New counterparties need research
        If ctx.NewCPDistinct > 0 Then
            WdComment doc, FirstPresent(docText, Array("OSDD", "Open-source", "open source", "counterpart")), _
                      "Newly identified counterparties in the New file (not in the Old file): " & ctx.NewCPDetail & _
                      ". Research them before filing."
        End If

        ' Rule from the New file not referenced in the narrative
        If ctx.RuleText <> "" Then
            For Each rulePart In Split(ctx.RuleText, "; ")
                If InStr(1, docText, CStr(rulePart), vbTextCompare) = 0 Then
                    If ruleMissing <> "" Then ruleMissing = ruleMissing & "; "
                    ruleMissing = ruleMissing & CStr(rulePart)
                End If
            Next rulePart
            If ruleMissing <> "" Then
                WdComment doc, FirstPresent(docText, Array("Type of Suspicious Activity")), _
                          "Rule triggered per the New file: " & ruleMissing & ". It is not referenced in the narrative."
            End If
        End If
    End If

    For Each f In flags
        WdComment doc, CStr(f(0)), CStr(f(1))
    Next f
End Sub

Private Function FirstPresent(ByVal docText As String, ByVal candidates As Variant) As String
    Dim c As Variant
    For Each c In candidates
        If CStr(c) <> "" Then
            If InStr(1, docText, CStr(c), vbTextCompare) > 0 Then
                FirstPresent = CStr(c)
                Exit Function
            End If
        End If
    Next c
End Function

' --- Text scanning helpers (positions are 1-based within the line text) ---

' End of the first sentence: the first ". " / ".<paragraph end>" not inside a number
Private Function SentenceEnd(ByVal t As String, ByVal fromPos As Long) As Long
    Dim i As Long, nextCh As String
    For i = fromPos To Len(t)
        If Mid$(t, i, 1) = "." Then
            nextCh = Mid$(t, i + 1, 1)
            If nextCh = " " Or nextCh = vbCr Or nextCh = "" Or nextCh = Chr(160) Then
                SentenceEnd = i
                Exit Function
            End If
        End If
    Next i
    SentenceEnd = Len(t)
End Function

' First m/d/yyyy or mm/dd/yyyy date in t between fromPos and toPos; returns its position (0 = none)
Private Function FindDateToken(ByVal t As String, ByVal fromPos As Long, ByVal toPos As Long, ByRef tokLen As Long) As Long
    Dim i As Long, L As Variant, tok As String, prevCh As String
    For i = fromPos To toPos
        If Mid$(t, i, 1) Like "#" Then
            If i > 1 Then prevCh = Mid$(t, i - 1, 1) Else prevCh = " "
            If Not (prevCh Like "[0-9/]") Then
                For Each L In Array(10, 9, 8)
                    tok = Mid$(t, i, L)
                    If tok Like "##/##/####" Or tok Like "#/##/####" Or tok Like "##/#/####" Or tok Like "#/#/####" Then
                        If Not (Mid$(t, i + L, 1) Like "#") Then
                            tokLen = L
                            FindDateToken = i
                            Exit Function
                        End If
                    End If
                Next L
            End If
        End If
    Next i
End Function

' First "$1,234.56" amount (with optional space after "$"); returns its position (0 = none)
Private Function FindAmountToken(ByVal t As String, ByVal fromPos As Long, ByVal toPos As Long, ByRef tokLen As Long) As Long
    Dim i As Long, j As Long
    For i = fromPos To toPos
        If Mid$(t, i, 1) = "$" Then
            j = i + 1
            Do While Mid$(t, j, 1) = " "
                j = j + 1
            Loop
            If Mid$(t, j, 1) Like "#" Then
                Do While Mid$(t, j, 1) Like "[0-9,.]"
                    j = j + 1
                Loop
                Do While j > i + 1
                    If Not (Mid$(t, j - 1, 1) Like "[.,]") Then Exit Do
                    j = j - 1
                Loop
                tokLen = j - i
                FindAmountToken = i
                Exit Function
            End If
        End If
    Next i
End Function

' First stand-alone number (digits/commas, not part of an amount or date) followed within
' three words by a word starting with one of followWords; returns its position (0 = none)
Private Function FindCountToken(ByVal t As String, ByVal fromPos As Long, ByVal toPos As Long, _
                                ByVal followWords As Variant, ByRef tokLen As Long) As Long
    Dim i As Long, j As Long, prevCh As String, nextCh As String, rest As String
    Dim words() As String, w As Variant, fw As Variant, n As Long

    i = fromPos
    Do While i <= toPos
        If Mid$(t, i, 1) Like "#" Then
            j = i
            Do While Mid$(t, j, 1) Like "[0-9,]"
                j = j + 1
            Loop
            Do While j > i + 1
                If Mid$(t, j - 1, 1) <> "," Then Exit Do
                j = j - 1
            Loop
            If i > 1 Then prevCh = Mid$(t, i - 1, 1) Else prevCh = " "
            nextCh = Mid$(t, j, 1)
            If Not (prevCh Like "[A-Za-z0-9$./]") And Not (nextCh Like "[A-Za-z0-9/.]") Then
                rest = LCase$(Mid$(t, j, 60))
                rest = Replace(Replace(Replace(rest, vbCr, " "), ",", " "), ".", " ")
                words = Split(Application.WorksheetFunction.Trim(rest), " ")
                n = 0
                For Each w In words
                    If CStr(w) <> "" Then
                        n = n + 1
                        For Each fw In followWords
                            If Left$(CStr(w), Len(CStr(fw))) = CStr(fw) Then
                                tokLen = j - i
                                FindCountToken = i
                                Exit Function
                            End If
                        Next fw
                        If n >= 3 Then Exit For
                    End If
                Next w
            End If
            i = j
        Else
            i = i + 1
        End If
    Loop
End Function

' --- Word helpers (late bound, no reference to the Word library needed) ---
Private Function GetWord() As Object
    If mWord Is Nothing Then
        On Error Resume Next
        Set mWord = GetObject(, "Word.Application")
        On Error GoTo 0
        If mWord Is Nothing Then
            Set mWord = CreateObject("Word.Application")
            mWordCreated = True
            On Error Resume Next
            mWord.Visible = False
            On Error GoTo 0
        Else
            mWordCreated = False
        End If
        On Error Resume Next
        mWordAlerts = mWord.DisplayAlerts
        mWord.DisplayAlerts = 0              ' wdAlertsNone
        On Error GoTo 0
    End If
    Set GetWord = mWord
End Function

' Restores Word; quits it only if this macro started it and no documents are open
Private Sub ReleaseWord()
    If mWord Is Nothing Then Exit Sub
    On Error Resume Next
    mWord.DisplayAlerts = mWordAlerts
    If mWordCreated Then
        If mWord.Documents.Count = 0 Then mWord.Quit SaveChanges:=0
    End If
    On Error GoTo 0
    Set mWord = Nothing
    mWordCreated = False
End Sub

Private Function WdFindIn(ByVal doc As Object, ByVal s As Long, ByVal e As Long, ByVal findText As String, _
                          ByVal wholeWord As Boolean) As Collection
    Dim col As New Collection
    Dim rng As Object
    Dim guard As Long

    Set WdFindIn = col
    If e > doc.Content.End Then e = doc.Content.End
    If findText = "" Or e <= s Then Exit Function
    Set rng = doc.Range(s, e)
    With rng.Find
        .ClearFormatting
        .Text = findText
        .Forward = True
        .Wrap = 0                            ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = wholeWord
        .MatchWildcards = False
    End With
    Do While rng.Find.Execute
        If rng.Start >= e Or rng.End > e Then Exit Do
        col.Add Array(rng.Start, rng.End)
        rng.Collapse 0                       ' wdCollapseEnd
        guard = guard + 1
        If guard > 1000 Then Exit Do
    Loop
End Function

Private Sub WdComment(ByVal doc As Object, ByVal anchorText As String, ByVal txt As String)
    Dim hits As Collection, h As Variant, rng As Object
    If anchorText <> "" Then
        Set hits = WdFindIn(doc, doc.Content.Start, doc.Content.End, anchorText, False)
        If hits.Count > 0 Then
            h = hits(1)
            Set rng = doc.Range(h(0), h(1))
        End If
    End If
    If rng Is Nothing Then Set rng = doc.Paragraphs(1).Range
    doc.Comments.Add rng, txt
End Sub

' --- Formatting helpers (independent of the PC's regional settings) ---
Private Function UsDate(ByVal v As Variant) As String
    If IsDate(v) Then UsDate = Format$(v, "mm\/dd\/yyyy")
End Function

Private Function UsAmount(ByVal x As Double) As String
    Dim s As String, sign As String
    If x < 0 Then sign = "-"
    s = Format$(Round(Abs(x), 2) * 100, "0")
    Do While Len(s) < 3
        s = "0" & s
    Loop
    UsAmount = sign & GroupThousands(Left$(s, Len(s) - 2)) & "." & Right$(s, 2)
End Function

Private Function UsNumber(ByVal n As Long) As String
    If n < 0 Then
        UsNumber = "-" & GroupThousands(CStr(-n))
    Else
        UsNumber = GroupThousands(CStr(n))
    End If
End Function

Private Function GroupThousands(ByVal digits As String) As String
    Dim out As String
    Do While Len(digits) > 3
        out = "," & Right$(digits, 3) & out
        digits = Left$(digits, Len(digits) - 3)
    Loop
    GroupThousands = digits & out
End Function

' --- Context for one alert, built from the Old/New comparison ---
Private Sub BuildNarrCtx(ByRef ctx As NarrCtx, ByVal ecmID As String, ByVal alertID As String, _
                         ByVal oldPath As String, ByVal newPath As String, _
                         ByRef oldSet As TxnSet, ByRef newSet As TxnSet, ByRef pr As PairResult)
    ctx.EcmID = ecmID
    ctx.AlertID = alertID
    ctx.OldFile = GetFileName(oldPath)
    ctx.NewFile = GetFileName(newPath)
    ctx.OldCount = oldSet.UniqueCount
    ctx.NewCount = newSet.UniqueCount
    ctx.OldTotal = oldSet.TotalAmt
    ctx.NewTotal = newSet.TotalAmt
    ctx.OldMin = oldSet.MinDate
    ctx.OldMax = oldSet.MaxDate
    ctx.NewMin = newSet.MinDate
    ctx.NewMax = newSet.MaxDate
    ctx.OldAlertCount = oldSet.AlertCount
    ctx.NewAlertCount = newSet.AlertCount
    ctx.OldAlertAmt = oldSet.AlertAmt
    ctx.NewAlertAmt = newSet.AlertAmt
    ctx.OldAlertMin = oldSet.AlertMin
    ctx.OldAlertMax = oldSet.AlertMax
    ctx.NewAlertMin = newSet.AlertMin
    ctx.NewAlertMax = newSet.AlertMax
    ctx.HasDrCr = oldSet.HasDrCr And newSet.HasDrCr
    ctx.OldCrCount = oldSet.CrCount
    ctx.OldCrAmt = oldSet.CrAmt
    ctx.OldDrCount = oldSet.DrCount
    ctx.OldDrAmt = oldSet.DrAmt
    ctx.NewCrCount = newSet.CrCount
    ctx.NewCrAmt = newSet.CrAmt
    ctx.NewDrCount = newSet.DrCount
    ctx.NewDrAmt = newSet.DrAmt
    FilePeriodDates oldPath, ctx.OldPeriodFrom, ctx.OldPeriodTo
    FilePeriodDates newPath, ctx.NewPeriodFrom, ctx.NewPeriodTo
    ctx.RuleText = newSet.RuleText
    ctx.AddedCount = pr.AddedCount
    ctx.AddedAmount = pr.AddedAmount
    ctx.AddedAlerted = pr.AddedAlerted
    ctx.AddedDetail = pr.AddedDetail
    ctx.DroppedCount = pr.DroppedCount
    ctx.NewCPDistinct = pr.NewCPDistinct
    ctx.NewCPDetail = pr.NewCPDetail
End Sub

' Period written in a file name ("06.01.2025 to 06.30.2026") as dates, or Empty
Private Sub FilePeriodDates(ByVal fPath As String, ByRef dFrom As Variant, ByRef dTo As Variant)
    Dim parts() As String, d As Date, ok As Boolean, txt As String
    dFrom = Empty
    dTo = Empty
    txt = FilePeriodText(fPath)
    If InStr(txt, " to ") = 0 Then Exit Sub
    parts = Split(txt, " to ")
    d = ParseDateValue(parts(0), ok)
    If Not ok Then Exit Sub
    dFrom = d
    d = ParseDateValue(parts(1), ok)
    If ok Then
        dTo = d
    Else
        dFrom = Empty
    End If
End Sub

Private Function AskNarrativeFolder() As String
    Dim folderPath As String
    If MsgBox("Also update the escalation narratives (Word) with the New file figures?" & vbCrLf & vbCrLf & _
              "Yes = select the narrative folder. Edited copies are saved in a new subfolder with each changed " & _
              "value highlighted in yellow; the original documents are not changed." & vbCrLf & _
              "No = compare transactions only.", vbQuestion + vbYesNo, "Update Narratives") <> vbYes Then Exit Function
    folderPath = Pick_Folder("Select the folder containing the escalation narratives (ECMID_ALERTID_*.docx)")
    If folderPath <> "" Then GrantMacAccess Array(folderPath)
    AskNarrativeFolder = folderPath
End Function

' Date value as mm/dd/yyyy, or the original text when it was not a readable date
Private Function DateText(ByVal v As Variant) As String
    If VarType(v) = vbDate Then
        DateText = UsDate(v)
    ElseIf IsEmpty(v) Then
        DateText = "(no date)"
    Else
        DateText = CStr(v)
    End If
End Function
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
    cm.Rule = GetFieldCol(hdr, 1, lastCol, FLD_RULE)
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
        Case FLD_RULE
            exacts = Array("alert information", "rule triggered", "rules triggered", "rule name", "rule", "scenario name", "scenario", "alert rule", "alert scenario")
            partials = Array("alert information", "scenario", "rule")
            excludes = Array("id", "date", "type", "count", "amount", "score")
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


Private Sub DictSet(ByVal d As Object, ByVal k As String, ByVal v As Variant)
    If TypeName(d) = "Dictionary" Then
        d(k) = v
    Else
        d.Remove k
        d.Add Array(k, v), k
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
