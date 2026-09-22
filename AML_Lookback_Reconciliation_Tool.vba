Attribute VB_Name = "AML_Period_Comparison"
' =========================================================================
' [AML] AML TRANSACTION MONITORING: LOOKBACK PERIOD BATCH COMPARISON COCKPIT
' =========================================================================
' Version: 6.5
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
'  5. Single-pair and batch-folder comparison. Windows, Excel 2016 or later.
'  6. Run_PreQC_Closure: pre-QC for closure (non-suspicious) alerts. Checks each
'     Alert Write-Up against its "Combined Alerted & Non Alerted Transactions"
'     file (alerted rows only).
'  7. Run_PreQC_Escalation: pre-QC for escalation narratives against the
'     lookback transaction file, using the same rules as the narrative updater.
'     Both pre-QC macros are read-only: they write a shaded review copy of each
'     narrative plus a summary and a findings sheet. Buttons on the dashboard.
'     The review copy also gets Word's spelling and grammar check: clear typing
'     slips are corrected there and listed, everything else is marked for the
'     reviewer (see PQ_PROOF).
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
' Narrative update: all changes are written into the narrative text itself.
' True also adds Word comments in the margin (Old vs New summary, new
' counterparties, rule not mentioned in the narrative).
Private Const NARRATIVE_ADD_COMMENTS As Boolean = False
' Date range written into narratives: "FILE" = the period in the New file's name
' (e.g. "05.01.2025 to 05.31.2026"), "DATA" = its first and last transaction date.
Private Const NARRATIVE_DATE_RANGE As String = "FILE"
' Note added as the first line of an updated narrative, highlighted yellow.
' Set to "" to add no note.
Private Const NARRATIVE_NOTE As String = "The data highlighted in yellow has been updated to reflect the revised " & _
    "date range and total transaction amount for this alert. QC reviewers are requested to confirm whether the " & _
    "in-scope transactions, timeline and amount align with the suspicious activity date range and the escalated " & _
    "scenarios within this case." 

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

Private Const FONT_UI As String = "Segoe UI"

' --- Sheet Names ---
Private Const SHEET_DASHBOARD As String = "Comparison Dashboard"
Private Const SHEET_ADDED_TXNS As String = "Discrepancy_Added_Txns"
Private Const SHEET_DROPPED_TXNS As String = "Discrepancy_Dropped_Txns"
Private Const SHEET_NEWCP_PIVOTS As String = "New_CP_Pivots"
Private Const SHEET_LEGACY_MODIFIED As String = "Discrepancy_Modified_Txns"   ' created by v3.0; removed on run

' --- Dashboard Layout ---
Private Const DASH_VERSION As String = "AML-PERIOD-COMPARISON-3.3"
Private Const DASH_TITLE As String = "AML TRANSACTION MONITORING - LOOKBACK PERIOD BATCH RECONCILIATION COCKPIT"
Private Const FIRST_DATA_ROW As Long = 15
' Zoom applied to the sheets after a run
Private Const DASH_ZOOM As Long = 55
Private Const SHEET_ZOOM As Long = 85
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
    MinAmt As Double           ' smallest / largest single transaction
    MaxAmt As Double
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
    CPBlank As Long            ' new transactions with no counterparty in the New file
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
    OldMinAmt As Double
    OldMaxAmt As Double
    NewMinAmt As Double
    NewMaxAmt As Double
    OldPeriodFrom As Variant
    OldPeriodTo As Variant
    NewPeriodFrom As Variant
    NewPeriodTo As Variant
    NarrFrom As Variant
    NarrTo As Variant
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

' --- Pre-QC (closure and escalation narratives) ---
' Sheets written by Run_PreQC_Closure and Run_PreQC_Escalation
Private Const SHEET_PQ_CLOSURE As String = "PreQC_Closure"
Private Const SHEET_PQ_CLOSURE_FIND As String = "PreQC_Closure_Findings"
Private Const SHEET_PQ_ESC As String = "PreQC_Escalation"
Private Const SHEET_PQ_ESC_FIND As String = "PreQC_Escalation_Findings"
' Shading written into the reviewed narrative copy. Only tokens the tool actually
' examined are shaded; untouched prose keeps its original background.
Private Const PQ_OK As Long = 14348502       ' soft mint  RGB(214, 240, 218) - verified
Private Const PQ_BAD As Long = 14079738      ' soft rose  RGB(250, 214, 214) - contradicts the file
Private Const PQ_EYE As Long = 11203327      ' yellow     RGB(255, 242, 170) - needs a human eye
Private Const PQ_NA As Long = 15788258       ' soft slate RGB(226, 232, 240) - not verifiable from the file
' Verdict labels used on the Pre-QC sheets
Private Const PQV_OK As String = "Verified"
Private Const PQV_BAD As String = "Mismatch"
Private Const PQV_EYE As String = "Check"
Private Const PQV_NA As String = "Not checked"
' A counterparty (or bank) absent from the narrative is reported when its share
' of the total reaches this fraction.
Private Const PQ_CP_SHARE As Double = 0.1
' Escalation wording checks: "all high dollar" is contradicted by transactions under
' PQ_SMALL_AMT; "consecutive days" when fewer than PQ_CLOSE_GAP_SHARE of the gaps between
' transaction dates are 1-2 days. Own-name transfers are reported between PQ_CP_SHARE
' and PQ_OWN_MAX_SHARE of the volume (above that the whole narrative is about them).
Private Const PQ_SMALL_AMT As Double = 10000
Private Const PQ_CLOSE_GAP_SHARE As Double = 0.25
Private Const PQ_OWN_MAX_SHARE As Double = 0.9
' Escalation narratives cover this many counterparties: the top one per alerted rule,
' then the largest in the lookback activity (see EscSelection)
Private Const PQ_ESC_CP_COUNT As Long = 5
' Output subfolder for the reviewed copies
Private Const PQ_OUT_PREFIX As String = "PreQC_Reviewed_"
' Word's spelling and grammar check on each reviewed copy (see ProofReviewedCopy).
' PQ_PROOF_FIX = False only marks and lists, without correcting anything.
' PQ_PROOF_MAX limits the spelling and grammar flags listed per narrative.
Private Const PQ_PROOF As Boolean = True
Private Const PQ_PROOF_FIX As Boolean = True
Private Const PQ_PROOF_MAX As Long = 40
Private Const PQV_PROOF As String = "Proofing"
' Words Word's dictionary may not know that are normal in these narratives
Private Const PQ_PROOF_WORDS As String = "ach iat dba aml kyc cdd edd osdd rfi sar cfsb ecm fintech " & _
    "ecommerce lookback onboard onboarded onboarding offboard offboarded offboarding omnichannel prefund " & _
    "prefunded prefunding payout payouts crypto cryptocurrency cryptocurrencies stablecoin stablecoins " & _
    "counterparty counterparties remitter remitters"
Private Const WD_UNDERLINE_WAVY As Long = 11
Private Const WD_COLOR_RED As Long = 255
Private Const WD_COLOR_BLUE As Long = 16711680

' Everything the closure checks need from one Combined Alerted & Non Alerted file
Private Type PreQCFacts
    EcmID As String
    AlertID As String
    AlertNum As String
    Customer As String
    RuleText As String
    Account As String
    AcctPrefix As String
    Ccy As String
    DrCr As String
    Instrument As String
    Program As String
    CustCountry As String
    AlertCount As Long
    AlertTotal As Double
    AlertMin As Variant
    AlertMax As Variant
    AllCount As Long
    AllTotal As Double
    NonCount As Long
    NonTotal As Double
    AlertCPs As Object
    NonCPs As Object
    CPAmt As Object
    SelfCPs As String
    HasData As Boolean
End Type

' Extra facts the escalation checks need from a lookback file, on top of the
' TxnSet the narrative updater uses (so both tools read the same numbers)
Private Type EscFacts
    AmtSet As Object       ' amount in cents -> label
    CntSet As Object       ' count -> label
    DateSet As Object      ' yyyymmdd -> label
    Groups As Object       ' "kind|value" -> Array(count, amount)
    CPs As Object          ' normalised counterparty -> display name
    CPAmt As Object        ' normalised counterparty -> amount
    Banks As Object        ' normalised bank name -> amount
    Programs As String     ' settlement program parties, "; " separated
    AcctPrefix As String
    Total As Double
    TxN As Long            ' transactions, one per unique transaction ID
    TxDate() As Date
    TxHasDate() As Boolean
    TxAmt() As Double
    TxCP() As String       ' normalised counterparty
    TxDir() As String      ' "C" / "D" / ""
    TxIns() As String      ' "ACH" / "WIRE" / ""
    TxAl() As Boolean      ' flagged "Is Alerted Transaction? = Yes"
    TxRule() As String     ' "Alert Information" (may hold several rules)
    Holders As Object      ' normalised account-holder name -> rows
    Holder As String       ' most frequent account holder
    Families As Object     ' first word shared by 2+ counterparties
    HasAlertFile As Boolean
    AlertFile As String
    AlertN As Long
    AlertTotal As Double
    AlertMin As Variant
    AlertMax As Variant
    AlertRule As String
    AlertMissing As Long   ' alerted rows absent from the lookback file
    LookbackAlerted As Long
    OnceFlags As String
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

    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, _
                                            FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
                                            Title:="Export AML Comparison Report to New Workbook")
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
    FinishAuditSheet wsAdded, 11, "No new transactions."
    FinishAuditSheet wsDropped, 9, "No dropped transactions."

    stepName = "Building New Counterparty Pivots"
    BuildNewCPPivotSheet wbDash, wsAdded, tot

    stepName = "Updating KPI Cards"
    UpdateKPICards wsDash, tot

    stepName = "Setting Sheet Zoom"
    SetSheetZoom wsAdded, SHEET_ZOOM
    SetSheetZoom wsDropped, SHEET_ZOOM
    Set ws = FindSheet(wbDash, SHEET_NEWCP_PIVOTS)
    If Not ws Is Nothing Then SetSheetZoom ws, SHEET_ZOOM
    SetSheetZoom wsDash, DASH_ZOOM

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
    tot.AlertRows.Add Array(ecmID, alertID, pr.AddedCount, pr.AddedAmount, pr.NewCPDistinct, pr.NewCPNames, pr.NewCPAgg, pr.CPBlank)
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
        If pr.CPBlank > 0 Then
            AddPart parts, "BLANK CP ON " & pr.CPBlank & " TXNS"
            If sev < 1 Then sev = 1
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

                ' Smallest / largest single transaction
                If Not IsEmpty(amt) Then
                    If ts.MaxAmt = 0 And ts.MinAmt = 0 Then
                        ts.MinAmt = Abs(CDbl(amt))
                        ts.MaxAmt = Abs(CDbl(amt))
                    Else
                        If Abs(CDbl(amt)) < ts.MinAmt Then ts.MinAmt = Abs(CDbl(amt))
                        If Abs(CDbl(amt)) > ts.MaxAmt Then ts.MaxAmt = Abs(CDbl(amt))
                    End If
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
        ReDim outArr(1 To n, 1 To 11)
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
            ElseIf cpFlag = "" Then
                pr.CPBlank = pr.CPBlank + 1
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
            outArr(r, 11) = ecmID & " | " & alertID

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
            ElseIf Left$(CStr(outArr(r, 10)), 7) = "Unknown" Or CStr(outArr(r, 10)) = "" Then
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
        CPStatus = ""                        ' no counterparty in the New file - left blank
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
    CreateActionButton ws, 12, "R", "R", "[PRE-QC] Closure Write-Ups", "Run_PreQC_Closure", RGB(109, 40, 217)
    CreateActionButton ws, 12, "S", "S", "[PRE-QC] Escalation Narratives", "Run_PreQC_Escalation", RGB(190, 24, 93)

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

    SetSheetZoom ws, DASH_ZOOM
    Set Setup_Comparison_Dashboard = ws
End Function

' Zoom is a window setting, so the sheet has to be activated briefly
Private Sub SetSheetZoom(ByVal ws As Worksheet, ByVal pct As Long)
    Dim prevSheet As Object
    On Error Resume Next
    Set prevSheet = ActiveSheet
    ws.Parent.Activate
    ws.Activate
    ActiveWindow.Zoom = pct
    If Not prevSheet Is Nothing Then prevSheet.Activate
    On Error GoTo 0
End Sub
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
        ws.Columns("L").ColumnWidth = 30 ' ECM / Alert (single filter for the pivot)
        lastColLetter = "L"
        headers = Array("ECM ID", "Alert ID", "Transaction ID", "Is Alerted?", "Transaction Date", "Amount", _
                        "Account No", "Transaction Description", "Counterparty Name", "New Counterparty?", "ECM / Alert")
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
' Two pivots on one sheet, both built on Discrepancy_Added_Txns and both driven
' by the same ECM ID / Alert ID slicers, so picking an alert shows only that
' alert's new transactions and only its newly identified counterparties:
'   1. NEW TRANSACTIONS      - date, transaction, counterparty and amount
'   2. NEW COUNTERPARTIES    - counterparty, number of transactions and amount
'                              (transactions with no counterparty included)
' If Excel cannot create a pivot, a plain per-alert table is written instead.
Private Sub BuildNewCPPivotSheet(ByVal wb As Workbook, ByVal wsAdded As Worksheet, ByRef tot As BatchTotals)
    Dim ws As Worksheet
    Dim item As Variant
    Dim nextRow As Long, lastAdded As Long, bottomRow As Long
    Dim totalAdded As Long, totalNewCP As Long, totalBlankCP As Long
    Dim srcAddr As String
    Dim pc As Object, ptTxn As Object, ptCP As Object

    Set ws = GetOrCreateWorksheet(wb, SHEET_NEWCP_PIVOTS)
    ClearPivotSheet ws

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 14

    ws.Rows(1).RowHeight = 28
    ws.Range("B1:G1").Merge
    With ws.Range("B1")
        .Value = "NEW TRANSACTIONS AND NEWLY IDENTIFIED COUNTERPARTIES - PICK AN ALERT WITH THE SLICERS"
        .Font.Name = FONT_UI
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_TXT
        .Interior.Color = COLOR_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    For Each item In tot.AlertRows
        totalAdded = totalAdded + item(2)
        totalNewCP = totalNewCP + item(4)
        totalBlankCP = totalBlankCP + item(7)
    Next item

    If totalAdded = 0 Then
        With ws.Cells(3, 2)
            .Value = "No new transactions in this run."
            .Font.Name = FONT_UI
            .Font.Italic = True
            .Font.Color = RGB(148, 163, 184)
        End With
        Exit Sub
    End If

    lastAdded = LastUsedRow(wsAdded)
    If lastAdded >= 4 Then
        srcAddr = "'" & wsAdded.Name & "'!" & _
                  wsAdded.Range(wsAdded.Cells(3, 2), wsAdded.Cells(lastAdded, 12)).Address(ReferenceStyle:=xlR1C1)
    End If
    If srcAddr = "" Then
        WriteStaticCPTable ws, 3, tot
        Exit Sub
    End If

    On Error Resume Next
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcAddr)
    On Error GoTo 0
    If pc Is Nothing Then
        WriteStaticCPTable ws, 3, tot
        Exit Sub
    End If

    ' --- 1. the selected alert's new transactions ---
    nextRow = 3
    WriteBlockTitle ws, nextRow, "NEW TRANSACTIONS (New file only) - " & UsNumber(totalAdded) & " in this run"
    bottomRow = TryCreateTxnPivot(pc, ws, nextRow + 6, "NewTxn_Pivot", ptTxn)
    If bottomRow = 0 Then
        WriteStaticCPTable ws, nextRow + 2, tot
        Exit Sub
    End If

    ' --- 2. the selected alert's newly identified counterparties ---
    nextRow = bottomRow + 3
    WriteBlockTitle ws, nextRow, "NEWLY IDENTIFIED COUNTERPARTIES - " & UsNumber(totalNewCP) & " in this run" & _
                    IIf(totalBlankCP > 0, "; " & UsNumber(totalBlankCP) & " new transaction(s) carry no counterparty", "")
    bottomRow = TryCreateNewCPPivot(pc, wb, ws, srcAddr, nextRow + 6, "NewCP_Pivot", ptCP)
    If bottomRow = 0 Then WriteStaticCPTable ws, nextRow + 2, tot

    ' --- slicers that drive both pivots ---
    AddCPSlicers wb, ws, ptTxn, ptCP, 3
End Sub

Private Sub WriteBlockTitle(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal txt As String)
    With ws.Cells(rowIdx, 2)
        .Value = txt
        .Font.Name = FONT_UI
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = COLOR_ALERT_RED
    End With
End Sub

' Transaction-level pivot: date / transaction / counterparty, amount as the value
Private Function TryCreateTxnPivot(ByVal pc As Object, ByVal ws As Worksheet, ByVal destRow As Long, _
                                   ByVal ptName As String, ByRef ptOut As Object) As Long
    Dim pt As Object, df As Object, fieldName As Variant

    On Error GoTo Fail
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Cells(destRow, 2), TableName:=ptName)
    pt.ManualUpdate = True

    With pt.PivotFields("ECM / Alert")
        .Orientation = xlPageField
        .Position = 1
    End With
    With pt.PivotFields("ECM ID")
        .Orientation = xlPageField
        .Position = 2
    End With
    With pt.PivotFields("Alert ID")
        .Orientation = xlPageField
        .Position = 3
    End With

    For Each fieldName In Array("Transaction Date", "Transaction ID", "Counterparty Name", "New Counterparty?")
        With pt.PivotFields(CStr(fieldName))
            .Orientation = xlRowField
            .Position = pt.RowFields.Count
            On Error Resume Next
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            .LabelRange.Ungroup                  ' keep real dates, not Excel's year/quarter groups
            On Error GoTo Fail
        End With
    Next fieldName

    Set df = pt.AddDataField(pt.PivotFields("Amount"), "Amount", xlSum)
    df.NumberFormat = "#,##0.00"

    On Error Resume Next
    pt.RowAxisLayout 1                           ' xlTabularRow
    pt.TableStyle2 = "PivotStyleLight16"
    On Error GoTo Fail
    pt.ManualUpdate = False

    Set ptOut = pt
    With pt.TableRange2
        TryCreateTxnPivot = .Row + .Rows.Count - 1
    End With
    Exit Function

Fail:
    Resume FailCleanup

FailCleanup:
    On Error Resume Next
    If Not pt Is Nothing Then pt.TableRange2.Clear
    TryCreateTxnPivot = 0
End Function

' Counterparty pivot: rows = counterparty, values = number of transactions and amount.
' "No" rows are hidden, so new counterparties and blank ones both show.
Private Function TryCreateNewCPPivot(ByVal pc As Object, ByVal wb As Workbook, ByVal ws As Worksheet, _
                                     ByVal srcAddr As String, ByVal destRow As Long, ByVal ptName As String, _
                                     ByRef ptOut As Object) As Long
    Dim pt As Object, df As Object

    On Error GoTo Fail
    Set pt = Nothing
    On Error Resume Next
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Cells(destRow, 2), TableName:=ptName)
    On Error GoTo Fail
    If pt Is Nothing Then                        ' some builds allow one pivot per cache only
        Set pt = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcAddr) _
                   .CreatePivotTable(TableDestination:=ws.Cells(destRow, 2), TableName:=ptName)
    End If

    pt.ManualUpdate = True
    With pt.PivotFields("ECM / Alert")
        .Orientation = xlPageField
        .Position = 1
    End With
    With pt.PivotFields("ECM ID")
        .Orientation = xlPageField
        .Position = 2
    End With
    With pt.PivotFields("Alert ID")
        .Orientation = xlPageField
        .Position = 3
    End With
    With pt.PivotFields("New Counterparty?")
        .Orientation = xlPageField
        .Position = 4
    End With
    With pt.PivotFields("Counterparty Name")
        .Orientation = xlRowField
        .Position = 1
    End With

    pt.AddDataField pt.PivotFields("Transaction ID"), "No of Trx", xlCount
    Set df = pt.AddDataField(pt.PivotFields("Amount"), "Sum of Transaction Amount", xlSum)
    df.NumberFormat = "#,##0.00"

    ShowNewAndBlankCPs pt
    pt.ManualUpdate = False

    On Error Resume Next
    pt.CompactLayoutRowHeader = "New Counterparty"
    pt.TableStyle2 = "PivotStyleLight16"
    pt.PivotFields("Counterparty Name").AutoSort xlDescending, "Sum of Transaction Amount"
    On Error GoTo Fail

    Set ptOut = pt
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

' ECM ID and Alert ID slicers, connected to both pivots. "Hide items with no
' data" makes them cascade: picking an ECM ID leaves only its alerts clickable.
Private Sub AddCPSlicers(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal ptTxn As Object, _
                         ByVal ptCP As Object, ByVal anchorRow As Long)
    Dim sc As Object
    Dim topPos As Double, leftPos As Double

    If ptTxn Is Nothing Then Exit Sub
    On Error GoTo Done
    topPos = ws.Cells(anchorRow, 8).Top
    leftPos = ws.Cells(anchorRow, 8).Left

    Set sc = wb.SlicerCaches.Add2(ptTxn, "ECM ID")
    sc.Slicers.Add ws, , "NewCP_ECM", "ECM ID", topPos, leftPos, 150, 190
    On Error Resume Next
    If Not ptCP Is Nothing Then sc.PivotTables.AddPivotTable ptCP
    sc.CrossFilterType = 3                       ' xlSlicerCrossFilterHideButtonsWithNoData
    On Error GoTo Done

    Set sc = wb.SlicerCaches.Add2(ptTxn, "Alert ID")
    sc.Slicers.Add ws, , "NewCP_Alert", "Alert ID", topPos, leftPos + 170, 190, 190
    On Error Resume Next
    If Not ptCP Is Nothing Then sc.PivotTables.AddPivotTable ptCP
    sc.CrossFilterType = 3
Done:
End Sub

' Hides only the "No" item, so new counterparties and blank ones both show
Private Sub ShowNewAndBlankCPs(ByVal pt As Object)
    Dim pf As Object, pi As Object
    Set pf = pt.PivotFields("New Counterparty?")
    On Error GoTo UseYesOnly
    pf.EnableMultiplePageItems = True
    For Each pi In pf.PivotItems
        pi.Visible = (StrComp(CStr(pi.Name), CP_NO, vbTextCompare) <> 0)
    Next pi
    Exit Sub
UseYesOnly:
    Resume Fallback
Fallback:
    On Error Resume Next
    pf.CurrentPage = CP_YES
End Sub

' Plain-table fallback with the same content as the pivots. Returns its last row.
Private Function WriteStaticCPTable(ByVal ws As Worksheet, ByVal startRow As Long, ByRef tot As BatchTotals) As Long
    Dim item As Variant, agg As Object, keys As Variant, v As Variant
    Dim i As Long, r As Long, totCount As Long
    Dim totAmt As Double

    WriteHeaderRow ws, startRow, 2, Array("ECM ID", "Alert ID", "New Counterparty", "No of Trx", "Sum of Transaction Amount")
    r = startRow + 1
    For Each item In tot.AlertRows
        Set agg = Nothing
        If IsObject(item(6)) Then Set agg = item(6)
        If Not agg Is Nothing Then
            keys = DictKeys(agg)
            For i = LBound(keys) To UBound(keys)
                v = DictGet(agg, CStr(keys(i)))
                ws.Cells(r, 2).NumberFormat = "@"
                ws.Cells(r, 3).NumberFormat = "@"
                ws.Cells(r, 4).NumberFormat = "@"
                ws.Cells(r, 2).Value = item(0)
                ws.Cells(r, 3).Value = item(1)
                ws.Cells(r, 4).Value = v(0)
                ws.Cells(r, 5).Value = v(1)
                ws.Cells(r, 6).Value = v(2)
                totCount = totCount + v(1)
                totAmt = totAmt + v(2)
                r = r + 1
            Next i
        End If
    Next item

    ws.Cells(r, 4).Value = "Grand Total"
    ws.Cells(r, 5).Value = totCount
    ws.Cells(r, 6).Value = totAmt
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 6)).Font.Bold = True

    With ws.Range(ws.Cells(startRow + 1, 2), ws.Cells(r, 6))
        .Font.Name = FONT_UI
        .Font.Size = 9
        With .Borders
            .LineStyle = xlContinuous
            .Color = COLOR_BORDER
            .Weight = xlThin
        End With
    End With
    ws.Range(ws.Cells(startRow + 1, 5), ws.Cells(r, 5)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(startRow + 1, 6), ws.Cells(r, 6)).NumberFormat = "#,##0.00"
    WriteStaticCPTable = r
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
    For i = ws.Shapes.Count To 1 Step -1     ' slicers from an earlier run
        ws.Shapes(i).Delete
    Next i
    For i = ws.PivotTables.Count To 1 Step -1
        ws.PivotTables(i).TableRange2.Clear
    Next i
    For i = ws.Parent.SlicerCaches.Count To 1 Step -1
        If ws.Parent.SlicerCaches(i).Slicers.Count = 0 Then ws.Parent.SlicerCaches(i).Delete
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
'     Lines 4-5: the date range and the first "N transactions totaling $X" pair
'     always become the New file figures (credit or debit figures when the
'     sentence says so, otherwise the overall ones); "amounts ranging from $A to
'     $B" becomes the New file's smallest and largest transaction. Later pairs in
'     the same sentence are only replaced when they equal an Old file figure.
'     K is increased by the number of newly identified counterparties. Any figure
'     replaced without matching the Old file is reported in the dashboard.
' Lines 1-3 take the New file's overall figures (count, total, date range).
' In the opening sentence and "Total Suspicious Dollar Amount" the total is
' rounded up to the next whole dollar ($109,206.78 -> $109,207); everywhere
' else amounts keep their cents.
' Alerted-only figures are never written anywhere in the narrative.
' Every date range written is the New file's own range: the period in its file
' name when there is one (NARRATIVE_DATE_RANGE = "FILE"), otherwise the first
' and last transaction date.
' Everywhere else in the document, any figure that equals an Old file figure
' (total, credit/debit total, transaction count, first/last date, file-name
' period) is replaced with the matching New figure, and "twelve-month review
' period" style phrases become the New file's period. Figures that match no Old
' figure are left alone.
' Every value is written into the narrative text itself, in the existing font
' (not tracked) and highlighted yellow. Nothing is added as a Word comment
' unless NARRATIVE_ADD_COMMENTS is True; anything that could not be updated is
' reported in the dashboard's "Narrative Update" column instead.
Private Function UpdateNarrative(ByRef ctx As NarrCtx, ByVal narrativeDir As String, ByRef tot As BatchTotals) As String
    Dim srcPath As String, outPath As String, note As String, stage As String, errDesc As String
    Dim doc As Object
    Dim edits As Collection, flags As Collection
    Dim copied As Boolean, wasTracking As Boolean
    Dim status As String, customer As String, flagText As String

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
    customer = CustomerFromFileName(srcPath)
    If customer = "" Then customer = CustomerFromDocument(srcPath)
    outPath = OutputNarrativePath(tot.NarrOutDir, srcPath, ctx.EcmID, ctx.AlertID, customer)
    On Error Resume Next
    FileCopy srcPath, outPath
    If Err.Number <> 0 Then          ' odd customer name - fall back to the plain name
        Err.Clear
        outPath = OutputNarrativePath(tot.NarrOutDir, srcPath, ctx.EcmID, ctx.AlertID, "")
        FileCopy srcPath, outPath
    End If
    On Error GoTo Fail
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
    If edits.Count > 0 Then InsertTopNote doc, NARRATIVE_NOTE
    If NARRATIVE_ADD_COMMENTS Then AddNarrativeComments doc, ctx, flags, edits.Count, note
    flagText = FlagSummary(flags)

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
    If flagText <> "" Then status = status & "; " & flagText
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

' The document as logical lines: paragraphs split on line breaks, because some
' narratives keep the whole SAR field block in a single paragraph.
' Each item is Array(start position, text).
Private Function LineSegments(ByVal doc As Object) As Collection
    Dim col As New Collection
    Dim p As Object
    Dim t As String, ch As String
    Dim base As Long, i As Long, segStart As Long

    For Each p In doc.Paragraphs
        t = p.Range.Text
        base = p.Range.Start
        segStart = 1
        For i = 1 To Len(t)
            ch = Mid$(t, i, 1)
            If ch = Chr(11) Or ch = vbCr Or ch = vbLf Then
                If i > segStart Then col.Add Array(base + segStart - 1, Mid$(t, segStart, i - segStart))
                segStart = i + 1
            End If
        Next i
        If Len(t) >= segStart Then col.Add Array(base + segStart - 1, Mid$(t, segStart))
    Next p
    Set LineSegments = col
End Function

' Customer name from the source file name (ECM_ALERT_<customer>_...)
Private Function CustomerFromFileName(ByVal srcPath As String) As String
    Dim tokens() As String, cand As String
    tokens = Split(StripExtension(GetFileName(srcPath)), "_")
    If UBound(tokens) >= 2 Then
        cand = Trim$(tokens(2))
        If Not IsNarrativeWord(cand) Then CustomerFromFileName = cand
    End If
End Function

' Customer name from the narrative's "Subject:" line (opens the source read-only)
Private Function CustomerFromDocument(ByVal srcPath As String) As String
    Dim doc As Object, sg As Variant
    Dim t As String, cand As String, pos As Long, n As Long

    On Error GoTo Done
    Set doc = GetWord().Documents.Open(FileName:=srcPath, ReadOnly:=True, AddToRecentFiles:=False)
    For Each sg In LineSegments(doc)
        t = CStr(sg(1))
        n = n + 1
        pos = InStr(1, t, "Subject", vbTextCompare)
        If pos > 0 Then
            pos = InStr(pos, t, ":")
            If pos > 0 Then
                cand = CleanCustomerName(Mid$(t, pos + 1))
                If cand <> "" Then
                    CustomerFromDocument = cand
                    Exit For
                End If
            End If
        End If
        If n >= 20 Then Exit For
    Next sg
Done:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=0
    On Error GoTo 0
End Function

' True for the descriptive parts of a file name ("Escalation Narrative", "Alert Write-Up", ...)
Private Function IsNarrativeWord(ByVal s As String) As Boolean
    Dim ls As String
    ls = LCase$(s)
    IsNarrativeWord = (ls = "" Or InStr(ls, "narrative") > 0 Or InStr(ls, "escalation") > 0 Or InStr(ls, "ecalation") > 0 _
                       Or InStr(ls, "write-up") > 0 Or InStr(ls, "write up") > 0 Or InStr(ls, "writeup") > 0 _
                       Or InStr(ls, "updated") > 0 Or InStr(ls, "alert write") > 0 _
                       Or InStr(ls, "combined") > 0 Or InStr(ls, "transactions") > 0)
End Function

' ECMID_ALERTID_Customer_Escalation Narrative.<ext> in the output folder
Private Function OutputNarrativePath(ByVal outDir As String, ByVal srcPath As String, ByVal ecmID As String, _
                                     ByVal alertID As String, ByVal customer As String) As String
    Dim fName As String, ext As String, newName As String, target As String, i As Long

    fName = GetFileName(srcPath)
    ext = Mid$(fName, InStrRev(fName, "."))
    newName = NarrativePrefix(ecmID, alertID)
    customer = SanitizeFileName(customer)
    If customer <> "" Then newName = newName & "_" & customer
    newName = newName & "_Escalation Narrative"

    ' keep the full path within the Windows limit
    Do While Len(outDir & newName & ext) > 240 And customer <> ""
        customer = Trim$(Left$(customer, Len(customer) - 5))
        newName = NarrativePrefix(ecmID, alertID)
        If customer <> "" Then newName = newName & "_" & customer
        newName = newName & "_Escalation Narrative"
    Loop

    target = outDir & newName & ext
    i = 1
    Do While FileExists(target)
        i = i + 1
        target = outDir & newName & " (" & i & ")" & ext
    Loop
    OutputNarrativePath = target
End Function

' Keeps just the name: stops at a line break, at the next numbered field
' ("2. Type of ...") or at a field colon, and caps the length
Private Function CleanCustomerName(ByVal s As String) As String
    Dim i As Long, ch As String, cut As Long

    For i = 1 To Len(s)
        If AscW(Mid$(s, i, 1)) < 32 Then
            s = Left$(s, i - 1)
            Exit For
        End If
    Next i

    cut = InStr(s, ":")
    If cut > 0 Then s = Left$(s, cut - 1)

    For i = 1 To Len(s) - 1
        If Mid$(s, i, 1) Like "#" And Mid$(s, i + 1, 1) = "." Then
            s = Left$(s, i - 1)
            Exit For
        End If
    Next i

    s = Trim$(s)
    Do While Len(s) > 0
        ch = Right$(s, 1)
        If ch <> "." And ch <> "," And ch <> "-" Then Exit Do
        s = Trim$(Left$(s, Len(s) - 1))
    Loop
    If Len(s) > 60 Then s = Trim$(Left$(s, 60))
    CleanCustomerName = s
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant, b As Variant
    Dim i As Long, out As String, ch As String

    ' drop every control character (line breaks, cell marks, ...)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If AscW(ch) >= 32 And AscW(ch) <> 127 Then out = out & ch
    Next i
    s = out

    bad = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
    For Each b In bad
        s = Replace(s, CStr(b), " ")
    Next b
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    Do While Len(s) > 0
        If Right$(s, 1) <> "." Then Exit Do
        s = Trim$(Left$(s, Len(s) - 1))
    Loop
    If Len(s) > 80 Then s = Trim$(Left$(s, 80))
    SanitizeFileName = s
End Function

' Everything that could not be updated, for the dashboard column
Private Function FlagSummary(ByVal flags As Collection) As String
    Dim f As Variant, s As String
    For Each f In flags
        If s <> "" Then s = s & " | "
        s = s & CStr(f(1))
    Next f
    If Len(s) > 500 Then s = Left$(s, 497) & "..."
    FlagSummary = s
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
    Dim segs As Collection, sg As Variant
    Dim t As String, lt As String
    Dim escS As Long, totS As Long, drS As Long, betS As Long, spcS As Long
    Dim escT As String, totT As String, drT As String, betT As String, spcT As String
    Dim newCount As String, newTotal As String
    Dim sentEnd As Long, k As Long, kLen As Long, kTxt As String, k2 As Long, kLen2 As Long, aPos As Long
    Dim pStart As Long
    Dim hCount As String, hAmount As String, hFrom As String, hTo As String, cv As Long

    escS = -1: totS = -1: drS = -1: betS = -1: spcS = -1
    Set segs = LineSegments(doc)
    For Each sg In segs
        t = CStr(sg(1))
        lt = LCase$(LTrim$(t))
        If escS < 0 And InStr(lt, "escalated") > 0 And InStr(lt, "totaling") > 0 Then
            escS = sg(0)
            escT = t
        ElseIf totS < 0 And InStr(lt, "total suspicious dollar amount") > 0 Then
            totS = sg(0)
            totT = t
        ElseIf drS < 0 And InStr(lt, "date range of suspicious activity") > 0 Then
            drS = sg(0)
            drT = t
        ElseIf betS < 0 And Left$(lt, 8) = "between " And InStr(lt, "totaling") > 0 Then
            betS = sg(0)
            betT = t
        ElseIf spcS < 0 And Left$(lt, 12) = "specifically" And InStr(lt, "between") > 0 And InStr(lt, "totaling") > 0 Then
            spcS = sg(0)
            spcT = t
        End If
    Next sg

    newCount = UsNumber(ctx.NewCount)
    newTotal = "$" & UsAmount(ctx.NewTotal)

    ' Lines 1-3: the New file's overall activity
    hCount = newCount
    hAmount = "$" & UsWholeAmount(ctx.NewTotal)
    hFrom = UsDate(ctx.NarrFrom)
    hTo = UsDate(ctx.NarrTo)

    If escS >= 0 Then
        sentEnd = SentenceEnd(escT, 1)
        k = NextCount(escT, 1, sentEnd, Array("transaction", "transfer", "txn"), kLen, cv)
        If k > 0 Then PlanFigure doc, edits, flags, escS, escT, k, kLen, HeaderPairs(ctx, "C"), hCount, _
                                 "opening sentence", "transaction count"
        aPos = InStr(1, escT, "totaling", vbTextCompare)
        If aPos < 1 Then aPos = 1
        k = FindAmountToken(escT, aPos, sentEnd, kLen)
        If k > 0 Then PlanFigure doc, edits, flags, escS, escT, k, kLen, HeaderPairs(ctx, "A"), hAmount, _
                                 "opening sentence", "amount"
    Else
        flags.Add Array("", "Opening sentence ('... to report N transactions totaling $X') not found; New file: " & _
                  hCount & " transactions totaling " & hAmount)
    End If

    If totS >= 0 Then
        k = FindAmountToken(totT, 1, Len(totT), kLen)
        If k > 0 Then PlanFigure doc, edits, flags, totS, totT, k, kLen, HeaderPairs(ctx, "A"), hAmount, _
                                 "Total Suspicious Dollar Amount", "amount"
    Else
        flags.Add Array("", "'Total Suspicious Dollar Amount' not found; New file: " & hAmount)
    End If

    If drS >= 0 Then
        k = FindDateToken(drT, 1, Len(drT), kLen)
        If k > 0 Then
            k2 = FindDateToken(drT, k + kLen, Len(drT), kLen2)
            PlanFigure doc, edits, flags, drS, drT, k, kLen, HeaderPairs(ctx, "F"), hFrom, _
                       "Date Range of Suspicious Activity", "start date"
            If k2 > 0 Then PlanFigure doc, edits, flags, drS, drT, k2, kLen2, HeaderPairs(ctx, "T"), hTo, _
                                      "Date Range of Suspicious Activity", "end date"
        End If
    Else
        flags.Add Array("", "'Date Range of Suspicious Activity' not found; New file: " & hFrom & " through " & hTo)
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
    ' Everywhere else: figures that came from the Old file
    For Each sg In segs
        pStart = sg(0)
        If pStart <> escS And pStart <> totS And pStart <> drS And pStart <> betS And pStart <> spcS Then
            SweepParagraph doc, ctx, edits, pStart, CStr(sg(1))
        End If
    Next sg
    PlanPeriodPhrase doc, ctx, edits

    If betS < 0 And spcS < 0 Then
        flags.Add Array("", "No 'Between [date] and [date], ... N transactions totaling $X' line was found. New file: " & _
                  newCount & " transactions totaling " & newTotal & " between " & UsDate(ctx.NarrFrom) & " and " & _
                  UsDate(ctx.NarrTo) & ".")
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

' Replaces Old file figures wherever else they appear: totals, credit/debit totals,
' transaction counts, the first/last transaction date inside a date range, and the
' Old file-name period inside a review-period sentence. Values that match no Old
' figure are left untouched (this is a sweep, not a checked line).
Private Sub SweepParagraph(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal edits As Collection, _
                           ByVal parStart As Long, ByVal t As String)
    Dim pEnd As Long, pos As Long, p As Long, n As Long, isPeriodText As Boolean
    Dim followWords As Variant

    pEnd = Len(t)
    followWords = Array("transaction", "transfer", "txn", "credit", "debit")
    isPeriodText = (InStr(1, t, "review period", vbTextCompare) > 0 Or InStr(1, t, "lookback", vbTextCompare) > 0 _
                    Or InStr(1, t, "look-back", vbTextCompare) > 0 Or InStr(1, t, "look back", vbTextCompare) > 0)

    ' amounts
    pos = 1
    Do
        p = FindAmountToken(t, pos, pEnd, n)
        If p = 0 Then Exit Do
        SwapIfOld doc, edits, parStart, t, p, n, AmountSweepPairs(ctx)
        pos = p + n
    Loop

    ' transaction counts
    pos = 1
    Do
        p = FindCountToken(t, pos, pEnd, followWords, n)
        If p = 0 Then Exit Do
        SwapIfOld doc, edits, parStart, t, p, n, CountSweepPairs(ctx)
        pos = p + n
    Loop

    ' dates that form a range
    pos = 1
    Do
        p = FindDateToken(t, pos, pEnd, n)
        If p = 0 Then Exit Do
        If IsRangeStartText(t, p + n) Then
            SwapIfOld doc, edits, parStart, t, p, n, Array(Array(UsDate(ctx.OldMin), UsDate(ctx.NarrFrom)), _
                                                           Array(UsDate(ctx.OldPeriodFrom), UsDate(ctx.NarrFrom)))
        ElseIf IsRangeEndText(t, p) Then
            SwapIfOld doc, edits, parStart, t, p, n, Array(Array(UsDate(ctx.OldMax), UsDate(ctx.NarrTo)), _
                                                           Array(UsDate(ctx.OldPeriodTo), UsDate(ctx.NarrTo)))
        End If
        pos = p + n
    Loop
End Sub

' Replaces the figure only when it equals one of the Old values
Private Sub SwapIfOld(ByVal doc As Object, ByVal edits As Collection, ByVal parStart As Long, ByVal t As String, _
                      ByVal pos As Long, ByVal n As Long, ByVal pairs As Variant)
    Dim cur As String, curKey As String, pr As Variant
    cur = Mid$(t, pos, n)
    curKey = NormValue(cur)
    For Each pr In pairs
        If CStr(pr(0)) <> "" And CStr(pr(1)) <> "" Then
            If NormValue(CStr(pr(0))) = curKey Then
                AddValueEdit doc, edits, parStart, t, pos, cur, InCountStyle(cur, CStr(pr(1)))
                Exit Sub
            End If
        End If
    Next pr
End Sub

Private Function AmountSweepPairs(ByRef ctx As NarrCtx) As Variant
    AmountSweepPairs = Array(Array("$" & UsAmount(ctx.OldTotal), "$" & UsAmount(ctx.NewTotal)), _
                             Array("$" & UsAmount(ctx.OldCrAmt), "$" & UsAmount(ctx.NewCrAmt)), _
                             Array("$" & UsAmount(ctx.OldDrAmt), "$" & UsAmount(ctx.NewDrAmt)))
End Function

Private Function CountSweepPairs(ByRef ctx As NarrCtx) As Variant
    CountSweepPairs = Array(Array(UsNumber(ctx.OldCount), UsNumber(ctx.NewCount)), _
                            Array(UsNumber(ctx.OldCrCount), UsNumber(ctx.NewCrCount)), _
                            Array(UsNumber(ctx.OldDrCount), UsNumber(ctx.NewDrCount)))
End Function

' "<word>-month review period" -> "review period from 06/01/2025 through 06/30/2026"
Private Sub PlanPeriodPhrase(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal edits As Collection)
    Dim hits As Collection, h As Variant, key As Variant
    Dim wordStart As Long

    If Not (IsDate(ctx.NewPeriodFrom) And IsDate(ctx.NewPeriodTo)) Then Exit Sub
    If IsDate(ctx.OldPeriodFrom) And IsDate(ctx.OldPeriodTo) Then
        If UsDate(ctx.OldPeriodFrom) = UsDate(ctx.NewPeriodFrom) And UsDate(ctx.OldPeriodTo) = UsDate(ctx.NewPeriodTo) Then Exit Sub
    End If

    For Each key In Array("month review period", "month lookback period", "month look-back period", "month look back period")
        Set hits = WdFindIn(doc, doc.Content.Start, doc.Content.End, CStr(key), False)
        For Each h In hits
            wordStart = WordStartBefore(doc, h(0))
            If wordStart >= 0 Then
                edits.Add Array(wordStart, h(1), doc.Range(wordStart, h(1)).Text, _
                                Mid$(CStr(key), 7) & " from " & UsDate(ctx.NewPeriodFrom) & " through " & UsDate(ctx.NewPeriodTo))
            End If
        Next h
    Next key
End Sub

' Start position of the word before "-month"/" month" (e.g. "twelve" in "twelve-month"), or -1
Private Function WordStartBefore(ByVal doc As Object, ByVal pos As Long) As Long
    Dim s As Long, t As String, i As Long, j As Long
    WordStartBefore = -1
    s = pos - 20
    If s < doc.Content.Start Then s = doc.Content.Start
    If pos <= s Then Exit Function
    t = doc.Range(s, pos).Text
    i = Len(t)
    If Not (Right$(t, 1) = "-" Or Right$(t, 1) = " ") Then Exit Function
    i = i - 1
    j = i
    Do While j > 0
        If Not (Mid$(t, j, 1) Like "[A-Za-z0-9]") Then Exit Do
        j = j - 1
    Loop
    If j = i Then Exit Function
    WordStartBefore = s + j
End Function

' Date at this position ends a range ("... and 06/02/2026")
Private Function IsRangeEndText(ByVal t As String, ByVal pos As Long) As Boolean
    Dim before As String
    If pos <= 1 Then Exit Function
    before = LCase$(Mid$(t, 1, pos - 1))
    IsRangeEndText = (before Like "*and " Or before Like "*through " Or before Like "*to " Or before Like "*until " _
                      Or before Like "*thru " Or before Like "*- " Or before Like "*-")
End Function

' Date ending at this position starts a range ("06/01/2025 and ...")
Private Function IsRangeStartText(ByVal t As String, ByVal pos As Long) As Boolean
    Dim after As String
    after = LCase$(Mid$(t, pos, 10))
    IsRangeStartText = (after Like " and*" Or after Like " through*" Or after Like " to *" Or after Like " until*" _
                        Or after Like " thru*" Or after Like " -*" Or after Like "-*")
End Function

' Old/New pair for a header figure, always the overall activity:
' "C" count, "A" amount, "F" first date, "T" last date
Private Function HeaderPairs(ByRef ctx As NarrCtx, ByVal what As String) As Variant
    Select Case what
        Case "C"
            HeaderPairs = Array(Array(UsNumber(ctx.OldCount), UsNumber(ctx.NewCount)))
        Case "A"
            ' the narrative may carry the Old total with or without cents
            HeaderPairs = Array(Array("$" & UsAmount(ctx.OldTotal), "$" & UsWholeAmount(ctx.NewTotal)), _
                                Array("$" & UsWholeAmount(ctx.OldTotal), "$" & UsWholeAmount(ctx.NewTotal)))
        Case "F"
            HeaderPairs = Array(Array(UsDate(ctx.OldMin), UsDate(ctx.NarrFrom)), _
                                Array(UsDate(ctx.OldPeriodFrom), UsDate(ctx.NarrFrom)))
        Case Else
            HeaderPairs = Array(Array(UsDate(ctx.OldMax), UsDate(ctx.NarrTo)), _
                                Array(UsDate(ctx.OldPeriodTo), UsDate(ctx.NarrTo)))
    End Select
End Function

' Dates, then every "N [credit/debit] transactions totaling $X" pair in the first sentence of a line.
' All or nothing: if any figure matches neither file the line is left unchanged and commented.
Private Sub PlanActivityLine(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal edits As Collection, ByVal flags As Collection, _
                             ByVal parStart As Long, ByVal t As String, ByVal lineName As String)
    Dim sentEnd As Long, p1 As Long, l1 As Long, p2 As Long, l2 As Long
    Dim pos As Long, p As Long, n As Long, nextP As Long, nn As Long, a As Long, an As Long, bound As Long
    Dim hint As String, followWords As Variant, cv As Long, cv2 As Long
    Dim firstPair As Boolean

    firstPair = True
    sentEnd = SentenceEnd(t, 1)
    followWords = Array("transaction", "transfer", "txn", "credit", "debit")

    p1 = FindDateToken(t, 1, sentEnd, l1)
    If p1 > 0 Then p2 = FindDateToken(t, p1 + l1, sentEnd, l2)
    If p1 > 0 And p2 > 0 Then
        PlanFigure doc, edits, flags, parStart, t, p1, l1, _
                   Array(Array(UsDate(ctx.OldMin), UsDate(ctx.NarrFrom)), Array(UsDate(ctx.OldPeriodFrom), UsDate(ctx.NarrFrom))), _
                   UsDate(ctx.NarrFrom), lineName, "start date"
        PlanFigure doc, edits, flags, parStart, t, p2, l2, _
                   Array(Array(UsDate(ctx.OldMax), UsDate(ctx.NarrTo)), Array(UsDate(ctx.OldPeriodTo), UsDate(ctx.NarrTo))), _
                   UsDate(ctx.NarrTo), lineName, "end date"
    Else
        flags.Add Array("", lineName & ": the date range was not found; New file range is " & UsDate(ctx.NarrFrom) & _
                  " to " & UsDate(ctx.NarrTo))
    End If

    pos = 1
    Do
        p = NextCount(t, pos, sentEnd, followWords, n, cv)
        If p = 0 Then Exit Do
        hint = CountDirection(t, p, n)
        If firstPair Then
            PlanFigure doc, edits, flags, parStart, t, p, n, CountPairs(ctx, hint), _
                       CStr(CountPairs(ctx, hint)(0)(1)), lineName, "transaction count"
        Else
            PlanFigure doc, edits, flags, parStart, t, p, n, CountPairs(ctx, hint), "", lineName, "transaction count"
        End If
        nextP = NextCount(t, p + n, sentEnd, followWords, nn, cv2)
        If nextP > 0 Then bound = nextP - 1 Else bound = sentEnd
        a = AmountAfterTotaling(t, p + n, bound, an)
        If a > 0 Then
            If firstPair Then
                PlanFigure doc, edits, flags, parStart, t, a, an, AmountPairs(ctx, hint), _
                           CStr(AmountPairs(ctx, hint)(0)(1)), lineName, "amount"
            Else
                PlanFigure doc, edits, flags, parStart, t, a, an, AmountPairs(ctx, hint), "", lineName, "amount"
            End If
        End If
        firstPair = False
        pos = p + n
    Loop

    ' "with amounts ranging from $671.28 to $50,000.00"
    PlanAmountRange doc, ctx, edits, flags, parStart, t, sentEnd, lineName

End Sub

' Replaces a figure with the New value of the first pair whose Old value it equals.
' Already-new figures are left alone. When nothing matches, forceNew (if given)
' is written anyway and the substitution is reported for checking.
Private Sub PlanFigure(ByVal doc As Object, ByVal edits As Collection, ByVal flags As Collection, ByVal parStart As Long, _
                       ByVal t As String, ByVal pos As Long, ByVal n As Long, ByVal pairs As Variant, _
                       ByVal forceNew As String, ByVal lineName As String, ByVal what As String)
    Dim cur As String, curKey As String, pr As Variant, oldList As String

    cur = Mid$(t, pos, n)
    If NumberWordValue(LCase$(cur)) > 0 Then
        curKey = NormValue(CStr(NumberWordValue(LCase$(cur))))
    Else
        curKey = NormValue(cur)
    End If
    For Each pr In pairs
        If CStr(pr(1)) <> "" Then
            If NormValue(CStr(pr(1))) = curKey Then Exit Sub
        End If
    Next pr
    For Each pr In pairs
        If CStr(pr(0)) <> "" And CStr(pr(1)) <> "" Then
            If NormValue(CStr(pr(0))) = curKey Then
                AddValueEdit doc, edits, parStart, t, pos, cur, InCountStyle(cur, CStr(pr(1)))
                Exit Sub
            End If
            If oldList <> "" Then oldList = oldList & " / "
            oldList = oldList & CStr(pr(0))
        End If
    Next pr

    If forceNew <> "" And NormValue(forceNew) <> curKey Then
        AddValueEdit doc, edits, parStart, t, pos, cur, InCountStyle(cur, forceNew)
        flags.Add Array("", lineName & ": " & what & " " & cur & " did not match the Old file (" & oldList & _
                  "); replaced with the New file figure " & forceNew & " - please check")
    ElseIf forceNew = "" Then
        flags.Add Array("", lineName & ": " & what & " " & cur & " matches neither the Old file (" & oldList & _
                  ") nor the New file; not changed")
    End If
End Sub

' "amounts ranging from $A to $B" -> smallest and largest transaction in the New file
Private Sub PlanAmountRange(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal edits As Collection, ByVal flags As Collection, _
                            ByVal parStart As Long, ByVal t As String, ByVal sentEnd As Long, ByVal lineName As String)
    Dim key As Variant, kp As Long, a1 As Long, n1 As Long, a2 As Long, n2 As Long

    If ctx.NewMaxAmt <= 0 Then Exit Sub
    For Each key In Array("ranging from", "ranging between", "ranged from", "ranged between", "range from")
        kp = InStr(1, t, CStr(key), vbTextCompare)
        If kp > 0 And kp < sentEnd Then
            a1 = FindAmountToken(t, kp, sentEnd, n1)
            If a1 > 0 Then a2 = FindAmountToken(t, a1 + n1, sentEnd, n2)
            If a1 > 0 And a2 > 0 Then
                PlanFigure doc, edits, flags, parStart, t, a1, n1, _
                           Array(Array("$" & UsAmount(ctx.OldMinAmt), "$" & UsAmount(ctx.NewMinAmt))), _
                           "$" & UsAmount(ctx.NewMinAmt), lineName, "smallest amount"
                PlanFigure doc, edits, flags, parStart, t, a2, n2, _
                           Array(Array("$" & UsAmount(ctx.OldMaxAmt), "$" & UsAmount(ctx.NewMaxAmt))), _
                           "$" & UsAmount(ctx.NewMaxAmt), lineName, "largest amount"
            End If
            Exit Sub
        End If
    Next key
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
    rest = LCase$(Replace(Mid$(t, pos + n, 40), Chr(160), " "))
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
        before = LCase$(Application.WorksheetFunction.Trim(Replace(Mid$(t, 1, pos - 1), Chr(160), " ")))
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
    s = TruePos(doc, parStart + offset - 1, Len(oldText))
    e = s + Len(oldText)
    If doc.Range(s, e).Text <> oldText Then
        Set hits = WdFindIn(doc, parStart, doc.Range(parStart, parStart).Paragraphs(1).Range.End, oldText, False)
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

' The new figure written the way the old one was: a count written as a word ("two") gets a
' word ("five", "Five" at the start of a sentence) up to twenty, otherwise digits
Private Function InCountStyle(ByVal oldText As String, ByVal newText As String) As String
    Dim words As Variant, v As Long, w As String
    InCountStyle = newText
    If NumberWordValue(LCase$(oldText)) = 0 Then Exit Function
    If newText Like "*[!0-9]*" Or newText = "" Then Exit Function
    v = CLng(newText)
    If v < 1 Or v > 20 Then Exit Function
    words = Array("one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", _
                  "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", _
                  "nineteen", "twenty")
    w = CStr(words(v - 1))
    If oldText = UCase$(oldText) Then
        w = UCase$(w)
    ElseIf Left$(oldText, 1) = UCase$(Left$(oldText, 1)) Then
        w = UCase$(Left$(w, 1)) & Mid$(w, 2)
    End If
    InCountStyle = w
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

' Adds the QC note as a new first paragraph, highlighted yellow, in the body font
Private Sub InsertTopNote(ByVal doc As Object, ByVal noteText As String)
    Dim rng As Object, para As Object, p As Object
    Dim fontName As String, fontSize As Single

    If Len(noteText) = 0 Then Exit Sub
    If InStr(1, doc.Paragraphs(1).Range.Text, Left$(noteText, 40), vbTextCompare) > 0 Then Exit Sub   ' already there

    For Each p In doc.Paragraphs
        If Len(p.Range.Text) > 80 Then
            fontName = p.Range.Font.Name
            fontSize = p.Range.Font.Size
            Exit For
        End If
    Next p

    doc.Paragraphs(1).Range.InsertParagraphBefore
    Set para = doc.Paragraphs(1)
    para.Range.InsertBefore noteText

    ' Highlight the paragraph that actually holds the note, without its paragraph mark
    Set para = doc.Paragraphs(1)
    Set rng = doc.Range(para.Range.Start, para.Range.End - 1)
    On Error Resume Next
    rng.Style = "Normal"
    On Error GoTo 0
    With rng.Font
        .Bold = False
        .Italic = False
        .Underline = 0                       ' wdUnderlineNone
        .Color = 0                           ' black
        If fontName <> "" Then .Name = fontName
        If fontSize > 0 Then .Size = fontSize
    End With
    rng.ParagraphFormat.Alignment = 0        ' wdAlignParagraphLeft
    rng.HighlightColorIndex = 7              ' wdYellow
    If rng.HighlightColorIndex <> 7 Then para.Range.HighlightColorIndex = 7
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
' End of the sentence starting at fromPos (the full stop, or Len(t)); "Inc.", "Ltd.", "Co."
' and initials do not end a sentence
Private Function SentenceEnd(ByVal t As String, ByVal fromPos As Long) As Long
    SentenceEnd = PQSentenceEnd(t, fromPos)
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
                rest = LCase$(Replace(Mid$(t, j, 60), Chr(160), " "))
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

' Amount rounded up to the next whole dollar, no cents ("109,207")
' For ordinary rounding instead, replace -Int(-x) with Round(x, 0).
Private Function UsWholeAmount(ByVal x As Double) As String
    Dim n As Double, sign As String
    If x < 0 Then sign = "-"
    n = -Int(-Abs(x))
    UsWholeAmount = sign & GroupThousands(Format$(n, "0"))
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
    ctx.OldMinAmt = oldSet.MinAmt
    ctx.OldMaxAmt = oldSet.MaxAmt
    ctx.NewMinAmt = newSet.MinAmt
    ctx.NewMaxAmt = newSet.MaxAmt
    ctx.NewCrCount = newSet.CrCount
    ctx.NewCrAmt = newSet.CrAmt
    ctx.NewDrCount = newSet.DrCount
    ctx.NewDrAmt = newSet.DrAmt
    FilePeriodDates oldPath, ctx.OldPeriodFrom, ctx.OldPeriodTo
    FilePeriodDates newPath, ctx.NewPeriodFrom, ctx.NewPeriodTo

    ' Date range written into the narrative: the New file's own period when its
    ' name carries one, otherwise its first and last transaction date
    If UCase$(NARRATIVE_DATE_RANGE) = "FILE" And IsDate(ctx.NewPeriodFrom) And IsDate(ctx.NewPeriodTo) Then
        ctx.NarrFrom = ctx.NewPeriodFrom
        ctx.NarrTo = ctx.NewPeriodTo
    Else
        ctx.NarrFrom = ctx.NewMin
        ctx.NarrTo = ctx.NewMax
    End If
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
' [HELPER] 11. FILES, FOLDERS AND WORKBOOKS
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
    Dim result As String, fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    fd.Title = promptTitle
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then result = fd.SelectedItems(1)
    result = Trim$(result)
    If result <> "" Then
        If Right$(result, 1) <> Application.PathSeparator Then result = result & Application.PathSeparator
    End If
    Pick_Folder = result
End Function

Private Function Pick_Excel_File(ByVal promptTitle As String) As String
    Dim vFile As Variant
    vFile = Application.GetOpenFilename("Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb", , promptTitle)
    If VarType(vFile) = vbBoolean Then
        Pick_Excel_File = ""
    ElseIf Not IsExcelFileName(CStr(vFile)) Then
        MsgBox "Please select an Excel file (.xlsx, .xlsm, .xls or .xlsb).", vbExclamation, "Not an Excel File"
        Pick_Excel_File = ""
    Else
        Pick_Excel_File = CStr(vFile)
    End If
End Function

Private Function ListExcelFiles(ByVal folderPath As String, ByVal excludePath As String) As Collection
    Dim col As New Collection
    Dim names As New Collection
    Dim sFile As String
    Dim arr As Variant, i As Long

    If Right$(folderPath, 1) <> Application.PathSeparator Then folderPath = folderPath & Application.PathSeparator

    ' list every file and filter by extension here
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

' Finds the period in a file name: "06.02.2025 to 06.03.2026" or,
' when that is absent, a compact pair such as "02012026 to 04012026"
Private Function FilePeriodText(ByVal fPath As String) As String
    Dim baseName As String, t1 As String, t2 As String

    baseName = StripExtension(GetFileName(fPath))
    CollectPeriodTokens baseName, False, t1, t2
    If t1 = "" Or t2 = "" Then CollectPeriodTokens baseName, True, t1, t2

    If t1 <> "" And t2 <> "" Then
        FilePeriodText = t1 & " to " & t2
    Else
        FilePeriodText = t1
    End If
End Function

' First two date tokens of a file name. compact = False looks for 06.02.2025 /
' 06-02-2025 / 06/02/2025, compact = True for 8-digit runs like 02012026
' (never one that follows a letter, so ALERT2365101 is not mistaken for a date).
Private Sub CollectPeriodTokens(ByVal baseName As String, ByVal compact As Boolean, ByRef t1 As String, ByRef t2 As String)
    Dim i As Long, j As Long
    Dim tok As String, prevCh As String
    Dim d As Date, ok As Boolean, isDate As Boolean
    Dim part As Variant

    t1 = ""
    t2 = ""
    i = 1
    Do While i <= Len(baseName)
        If Mid$(baseName, i, 1) Like "#" Then
            j = i
            Do While j <= Len(baseName)
                If compact Then
                    If Not (Mid$(baseName, j, 1) Like "#") Then Exit Do
                Else
                    If Not (Mid$(baseName, j, 1) Like "[0-9./-]") Then Exit Do
                End If
                j = j + 1
            Loop
            tok = Mid$(baseName, i, j - i)
            Do While Len(tok) > 0
                If InStr("./-", Right$(tok, 1)) = 0 Then Exit Do
                tok = Left$(tok, Len(tok) - 1)
            Loop

            If i > 1 Then prevCh = Mid$(baseName, i - 1, 1) Else prevCh = " "
            isDate = False
            If compact Then
                If Len(tok) = 8 And Not (prevCh Like "[A-Za-z]") Then
                    d = ParseDateValue(tok, ok)
                    isDate = ok
                End If
            Else
                isDate = LooksLikeDateToken(tok)
            End If

            If isDate Then
                If t1 = "" Then
                    t1 = tok
                ElseIf t2 = "" Then
                    t2 = tok
                End If
            ElseIf Not compact Then
                ' two dates joined by a hyphen, e.g. "6.1.2025-6.3.2026"
                If InStr(tok, "-") > 0 Then
                    For Each part In Split(tok, "-")
                        If LooksLikeDateToken(CStr(part)) Then
                            If t1 = "" Then
                                t1 = CStr(part)
                            ElseIf t2 = "" Then
                                t2 = CStr(part)
                            End If
                        End If
                    Next part
                End If
            End If
            i = j
        Else
            i = i + 1
        End If
    Loop
End Sub

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

    ' Compact date without separators: 02012026 (per TEXT_DATE_ORDER) or 20260201
    If IsAllDigits(s) And Len(s) = 8 Then
        y = CLng(Left$(s, 4))
        If y >= 1950 And y <= 2100 Then
            If BuildDate(y, CLng(Mid$(s, 5, 2)), CLng(Mid$(s, 7, 2)), result) Then
                ParseDateValue = result
                ok = True
                Exit Function
            End If
        End If
        y = CLng(Mid$(s, 5, 4))
        If TEXT_DATE_ORDER = "DMY" Then
            d = CLng(Left$(s, 2))
            m = CLng(Mid$(s, 3, 2))
        Else
            m = CLng(Left$(s, 2))
            d = CLng(Mid$(s, 3, 2))
        End If
        If m > 12 And d <= 12 Then
            tmp = m
            m = d
            d = tmp
        End If
        If BuildDate(y, m, d, result) Then
            ParseDateValue = result
            ok = True
        End If
        Exit Function
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
' [DICT] 14. KEY-VALUE LOOKUP (Scripting.Dictionary, keys case-insensitive)
' =========================================================================
Private Function CreateLookupDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' vbTextCompare
    Set CreateLookupDict = d
End Function

Private Function DictExists(ByVal d As Object, ByVal k As String) As Boolean
    DictExists = d.Exists(k)
End Function

Private Sub DictAdd(ByVal d As Object, ByVal k As String, ByVal v As Variant)
    d.Add k, v
End Sub

Private Sub DictSet(ByVal d As Object, ByVal k As String, ByVal v As Variant)
    d(k) = v
End Sub

Private Function DictGet(ByVal d As Object, ByVal k As String) As Variant
    DictGet = d(k)
End Function

Private Function DictCount(ByVal d As Object) As Long
    DictCount = d.Count
End Function

Private Function DictKeys(ByVal d As Object) As Variant
    DictKeys = d.Keys
End Function

' =========================================================================
' [PQC] PRE-QC: SHARED RUNNER PIECES
' =========================================================================
' Both pre-QC macros check every hardcoded figure in a narrative against its
' transaction file. The original narrative is never changed: a reviewed copy
' goes to a new subfolder with each examined token shaded.
'   green  = matches the file        red   = contradicts the file
'   yellow = needs a reviewer's eye  grey  = examined, not verifiable from the file
' Unshaded text was not checked.
' =========================================================================

' Workbook the report sheets go to: the one holding the dashboard, like the comparison
Private Function PreQCBook() As Workbook
    Set PreQCBook = ActiveWorkbook
    If PreQCBook Is Nothing Then
        MsgBox "Open the workbook that should hold the pre-QC sheets first.", vbExclamation, "Pre-QC"
    End If
End Function

' "<folder>PreQC_Reviewed_<timestamp><sep>", or the folder itself if it cannot be created
Private Function PreQCOutDir(ByVal baseDir As String) As String
    Dim outDir As String
    outDir = baseDir & PQ_OUT_PREFIX & Format(Now, "yyyymmdd_hhnnss")
    On Error Resume Next
    MkDir outDir
    On Error GoTo 0
    If FolderExists(outDir) Then
        PreQCOutDir = outDir & Application.PathSeparator
    Else
        PreQCOutDir = baseDir
    End If
End Function

Private Sub PreQCDoneMessage(ByVal nDocs As Long, ByVal okTot As Long, ByVal badTot As Long, _
                             ByVal eyeTot As Long, ByVal outDir As String, ByVal kind As String, _
                             ByVal fixTot As Long, ByVal revTot As Long)
    Dim msg As String
    msg = kind & " pre-QC complete." & vbCrLf & vbCrLf & _
          nDocs & " alert(s) checked" & vbCrLf & _
          okTot & " verified, " & badTot & " mismatch, " & eyeTot & " to look at" & vbCrLf
    If PQ_PROOF Then msg = msg & "Spelling / grammar: " & fixTot & " corrected, " & revTot & " to review" & vbCrLf
    msg = msg & vbCrLf & "Reviewed copies: " & outDir
    If badTot > 0 Then
        MsgBox msg, vbExclamation, "Pre-QC"
    Else
        MsgBox msg, vbInformation, "Pre-QC"
    End If
End Sub

' =========================================================================
' [PQC] PRE-QC FOR CLOSURE (NON-SUSPICIOUS) ALERTS
' =========================================================================
' Reads each "<ECM>_<ALERT>_Combined Alerted & Non Alerted Transactions.xlsx"
' with its "<ECM>_<ALERT>_<Customer>_Alert Write-Up.docx". Closure write-ups
' state the ALERTED transactions only, so the figures are checked against the
' rows flagged "Is Alerted Transaction? = Yes".
' =========================================================================
Public Sub Run_PreQC_Closure()
    Dim st As AppState
    Dim srcDir As String, outDir As String
    Dim jobs As Collection, job As Variant
    Dim wsDash As Worksheet, wsFind As Worksheet
    Dim fx As PreQCFacts
    Dim items As Collection
    Dim dashRow As Long, findRow As Long
    Dim nOK As Long, nBad As Long, nEye As Long
    Dim okTot As Long, badTot As Long, eyeTot As Long, nDocs As Long
    Dim nFix As Long, nRev As Long, fixTot As Long, revTot As Long
    Dim outPath As String, msg As String
    Dim wbOut As Workbook

    Set wbOut = PreQCBook()
    If wbOut Is Nothing Then Exit Sub
    srcDir = Pick_Folder("Select the folder with the Combined transaction files and Alert Write-Ups")
    If srcDir = "" Then Exit Sub

    Set jobs = PreQCClosureJobs(srcDir)
    If jobs.Count = 0 Then
        MsgBox "No '...Combined Alerted & Non Alerted Transactions' file was found in that folder.", _
               vbExclamation, "Pre-QC"
        Exit Sub
    End If
    outDir = PreQCOutDir(srcDir)

    SaveAndSpeedUp st
    On Error GoTo Failed

    Set wsDash = GetOrCreateWorksheet(wbOut, SHEET_PQ_CLOSURE)
    Set wsFind = GetOrCreateWorksheet(wbOut, SHEET_PQ_CLOSURE_FIND)
    PreQCPrepareSheets wsDash, wsFind, "PRE-QC: CLOSURE (NON-SUSPICIOUS) ALERT WRITE-UPS", _
                       Array("ECM ID", "Alert ID", "Customer", "Rule", "Alerted Txns", _
                             "Alerted Amount", "Alerted Period", "Verified", "Mismatch", _
                             "To Check", "Status", "Write-Up", "Reviewed Copy", "Spelling / Grammar")
    dashRow = 4
    findRow = 4

    For Each job In jobs
        nDocs = nDocs + 1
        ClearFacts fx
        Set items = New Collection
        outPath = ""
        outPath = RunClosureJob(job, outDir, fx, items)
        CountVerdicts items, nOK, nBad, nEye
        okTot = okTot + nOK
        badTot = badTot + nBad
        eyeTot = eyeTot + nEye
        WritePreQCDash wsDash, dashRow, fx, nOK, nBad, nEye, CStr(job(1)), outPath
        wsDash.Cells(dashRow, 15).Value = ProofSummary(items)
        ProofCounts items, nFix, nRev
        fixTot = fixTot + nFix
        revTot = revTot + nRev
        WritePreQCFindings wsFind, findRow, CStr(job(2)), CStr(job(3)), items
        dashRow = dashRow + 1
    Next job

    FinishPreQCSheet wsDash, 14, dashRow - 1, "No alerts were checked."
    FinishPreQCSheet wsFind, 7, findRow - 1, "Nothing flagged - every checked figure matched the file."
    ReleaseWord
    RestoreApp st
    wsDash.Activate
    PreQCDoneMessage nDocs, okTot, badTot, eyeTot, outDir, "Closure", fixTot, revTot
    Exit Sub

Failed:
    msg = Err.Description
    ReleaseWord
    RestoreApp st
    MsgBox "Pre-QC stopped: " & msg, vbCritical, "Pre-QC"
End Sub

' One alert; a file that cannot be read is reported on its row instead of stopping the run
Private Function RunClosureJob(ByVal job As Variant, ByVal outDir As String, ByRef fx As PreQCFacts, _
                               ByVal items As Collection) As String
    Dim errDesc As String
    On Error GoTo Failed
    LoadClosureFacts CStr(job(0)), CStr(job(1)), CStr(job(2)), CStr(job(3)), fx
    If CStr(job(1)) = "" Then
        AddPQItem items, "Write-up", "", "", PQV_NA, "no Alert Write-Up found for this alert"
    Else
        RunClosureJob = ReviewWriteUp(CStr(job(1)), outDir, fx, items)
    End If
    Exit Function
Failed:
    errDesc = Err.Description
    fx.EcmID = CStr(job(2))
    fx.AlertID = CStr(job(3))
    AddPQItem items, "Transaction file", GetFileName(CStr(job(0))), "", PQV_EYE, "could not be read (" & errDesc & ")"
End Function

' Pairs every Combined file in the folder with its Alert Write-Up
Private Function PreQCClosureJobs(ByVal folderPath As String) As Collection
    Dim jobs As New Collection
    Dim files As Collection, f As Variant
    Dim fName As String, ecmID As String, alertID As String, matchKey As String, note As String

    Set files = ListExcelFiles(folderPath, "")
    For Each f In files
        fName = GetFileName(CStr(f))
        If InStr(1, fName, "Combined", vbTextCompare) > 0 Then
            ExtractFileIDs CStr(f), ecmID, alertID, matchKey
            If ecmID <> "" Then
                jobs.Add Array(CStr(f), FindNarrativeFile(folderPath, ecmID, alertID, note), ecmID, alertID)
            End If
        End If
    Next f
    Set PreQCClosureJobs = jobs
End Function

Private Sub ClearFacts(ByRef fx As PreQCFacts)
    Dim blank As PreQCFacts
    fx = blank
    Set fx.AlertCPs = CreateLookupDict()
    Set fx.NonCPs = CreateLookupDict()
    Set fx.CPAmt = CreateLookupDict()
End Sub

' =========================================================================
' [PQC] READING THE COMBINED FILE
' =========================================================================
Private Sub LoadClosureFacts(ByVal fPath As String, ByVal docPath As String, ByVal ecmID As String, _
                             ByVal alertID As String, ByRef fx As PreQCFacts)
    Dim wb As Workbook, ws As Worksheet
    Dim openedByUs As Boolean

    fx.EcmID = ecmID
    fx.AlertID = alertID
    If docPath <> "" Then fx.Customer = CustomerFromFileName(docPath)

    Set wb = GetOrOpenWorkbook(fPath, openedByUs)
    On Error GoTo Cleanup

    Set ws = FindSheet(wb, "Raw Transactions")
    If ws Is Nothing Then Set ws = FindTransactionSheet(wb)
    If Not ws Is Nothing Then ReadCombinedSheet ws, fx

    ' "Non Alerted Transaction" is a curated review-window subset, not every
    ' non-alerted row of the raw sheet, so it is read separately.
    Set ws = FindSheet(wb, "Non Alerted Transaction")
    If Not ws Is Nothing Then ReadNonAlertedSheet ws, fx

Cleanup:
    On Error Resume Next
    If openedByUs Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub ReadCombinedSheet(ByVal ws As Worksheet, ByRef fx As PreQCFacts)
    Dim block As Variant
    Dim lastRow As Long, lastCol As Long, hdrRow As Long
    Dim cAlerted As Long, cAmt As Long, cTxDate As Long, cCP As Long, cDrCr As Long
    Dim cNum As Long, cRule As Long, cAcct As Long, cCcy As Long, cCode As Long
    Dim cPType As Long, cPName As Long, cOCty As Long, cBCty As Long
    Dim r As Long, isAlert As Boolean, okDate As Boolean
    Dim amt As Double, dv As Date, cp As String, key As String, dr As String
    Dim custKey As String

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedCol(ws)
    If lastRow < 2 Then Exit Sub
    If lastCol < 1 Then Exit Sub

    hdrRow = CombinedHeaderRow(ws, lastCol)
    If hdrRow = 0 Then Exit Sub
    If hdrRow >= lastRow Then Exit Sub

    block = ReadBlock(ws, hdrRow, 1, lastRow, lastCol)

    cAlerted = NamedCol(block, lastCol, "Is Alerted Transaction?")
    cAmt = NamedCol(block, lastCol, "Transaction Amount")
    cTxDate = NamedCol(block, lastCol, "Transaction Date")
    cCP = NamedCol(block, lastCol, "Counterparty")
    cDrCr = NamedCol(block, lastCol, "Dr Cr")
    cNum = NamedCol(block, lastCol, "Velocity Alert ID")
    cRule = NamedCol(block, lastCol, "Alert Information")
    cAcct = NamedCol(block, lastCol, "Account No")
    cCcy = NamedCol(block, lastCol, "Currency")
    cCode = NamedCol(block, lastCol, "Transaction Code")
    cPType = NamedCol(block, lastCol, "Party Type2")
    cPName = NamedCol(block, lastCol, "Party Name2")
    cOCty = NamedCol(block, lastCol, "Originator Country")
    cBCty = NamedCol(block, lastCol, "Beneficiary Country")
    If cAmt = 0 Then Exit Sub

    custKey = NormCP(fx.Customer)

    For r = 2 To UBound(block, 1)
        If Not IsEmptyBlockRow(block, r, lastCol) Then
            amt = ValueToNumber(BlockCell(block, r, cAmt))
            dv = ParseDateValue(BlockCell(block, r, cTxDate), okDate)
            cp = Trim$(CellText(BlockCell(block, r, cCP)))
            dr = UCase$(Trim$(CellText(BlockCell(block, r, cDrCr))))
            isAlert = (LCase$(Trim$(CellText(BlockCell(block, r, cAlerted)))) = "yes")

            fx.AllCount = fx.AllCount + 1
            fx.AllTotal = fx.AllTotal + amt

            If isAlert Then
                fx.HasData = True
                fx.AlertCount = fx.AlertCount + 1
                fx.AlertTotal = fx.AlertTotal + amt
                If okDate Then
                    If Not IsDate(fx.AlertMin) Then
                        fx.AlertMin = dv
                        fx.AlertMax = dv
                    Else
                        If dv < CDate(fx.AlertMin) Then fx.AlertMin = dv
                        If dv > CDate(fx.AlertMax) Then fx.AlertMax = dv
                    End If
                End If
                If fx.DrCr = "" Then fx.DrCr = dr
                If fx.AlertNum = "" Then fx.AlertNum = Trim$(CellText(BlockCell(block, r, cNum)))
                If fx.RuleText = "" Then fx.RuleText = Trim$(CellText(BlockCell(block, r, cRule)))
                If fx.Account = "" Then fx.Account = Trim$(CellText(BlockCell(block, r, cAcct)))
                If fx.Ccy = "" Then fx.Ccy = Trim$(CellText(BlockCell(block, r, cCcy)))
                If fx.Instrument = "" Then fx.Instrument = InstrumentOf(CellText(BlockCell(block, r, cCode)))
                If fx.CustCountry = "" Then
                    If dr = "CR" Then
                        fx.CustCountry = Trim$(CellText(BlockCell(block, r, cBCty)))
                    Else
                        fx.CustCountry = Trim$(CellText(BlockCell(block, r, cOCty)))
                    End If
                End If
                If cp <> "" Then
                    key = NormCP(cp)
                    If key = custKey Then
                        AddToList fx.SelfCPs, cp
                    Else
                        If Not DictExists(fx.AlertCPs, key) Then DictAdd fx.AlertCPs, key, cp
                        If DictExists(fx.CPAmt, key) Then
                            DictSet fx.CPAmt, key, CDbl(DictGet(fx.CPAmt, key)) + amt
                        Else
                            DictAdd fx.CPAmt, key, amt
                        End If
                    End If
                End If
            End If

            If LCase$(Trim$(CellText(BlockCell(block, r, cPType)))) = "settlementprogram" Then
                AddToList fx.Program, Trim$(CellText(BlockCell(block, r, cPName)))
            End If
        End If
    Next r

    If fx.Account <> "" Then
        If InStr(fx.Account, "-") > 0 Then
            fx.AcctPrefix = Left$(fx.Account, InStr(fx.Account, "-") - 1)
        Else
            fx.AcctPrefix = fx.Account
        End If
    End If
End Sub

Private Sub ReadNonAlertedSheet(ByVal ws As Worksheet, ByRef fx As PreQCFacts)
    Dim block As Variant
    Dim lastRow As Long, lastCol As Long, hdrRow As Long
    Dim cAmt As Long, cCP As Long
    Dim r As Long, cp As String, key As String

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedCol(ws)
    If lastRow < 2 Then Exit Sub
    If lastCol < 1 Then Exit Sub

    hdrRow = CombinedHeaderRow(ws, lastCol)
    If hdrRow = 0 Then Exit Sub
    If hdrRow >= lastRow Then Exit Sub

    block = ReadBlock(ws, hdrRow, 1, lastRow, lastCol)
    cAmt = NamedCol(block, lastCol, "Transaction Amount")
    cCP = NamedCol(block, lastCol, "Counterparty")
    If cAmt = 0 Then Exit Sub

    For r = 2 To UBound(block, 1)
        If Not IsEmptyBlockRow(block, r, lastCol) Then
            fx.NonCount = fx.NonCount + 1
            fx.NonTotal = fx.NonTotal + ValueToNumber(BlockCell(block, r, cAmt))
            cp = Trim$(CellText(BlockCell(block, r, cCP)))
            If cp <> "" Then
                key = NormCP(cp)
                If Not DictExists(fx.NonCPs, key) Then DictAdd fx.NonCPs, key, cp
            End If
        End If
    Next r
End Sub

' Row holding "Transaction Amount" within the first HEADER_SCAN_ROWS rows
Private Function CombinedHeaderRow(ByVal ws As Worksheet, ByVal lastCol As Long) As Long
    Dim r As Long, c As Long, h As String
    For r = 1 To HEADER_SCAN_ROWS
        For c = 1 To lastCol
            h = LCase$(Trim$(CellText(ws.Cells(r, c).Value)))
            If h = "transaction amount" Then
                CombinedHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
End Function

' Column index of an exact header name in row 1 of the block (0 = absent). The block is
' passed ByRef: a Variant array passed ByVal is copied on every call.
Private Function NamedCol(ByRef block As Variant, ByVal lastCol As Long, ByVal colName As String) As Long
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim$(CellText(BlockCell(block, 1, c))), colName, vbTextCompare) = 0 Then
            NamedCol = c
            Exit Function
        End If
    Next c
End Function

Private Function BlockCell(ByRef block As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c < 1 Then Exit Function
    If r < 1 Then Exit Function
    If r > UBound(block, 1) Then Exit Function
    If c > UBound(block, 2) Then Exit Function
    BlockCell = block(r, c)
End Function

Private Function IsEmptyBlockRow(ByRef block As Variant, ByVal r As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long
    For c = 1 To lastCol
        If Trim$(CellText(BlockCell(block, r, c))) <> "" Then Exit Function
    Next c
    IsEmptyBlockRow = True
End Function

Private Function ValueToNumber(ByVal v As Variant) As Double
    Dim s As String, i As Long, ch As String, out As String
    If IsNumeric(v) Then
        ValueToNumber = CDbl(v)
        Exit Function
    End If
    s = CellText(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9.]" Then out = out & ch
    Next i
    If out <> "" Then
        If IsNumeric(out) Then ValueToNumber = CDbl(out)
    End If
End Function

' "ACH" or "WIRE" from the transaction code; "" when neither is recognised
Private Function InstrumentOf(ByVal code As String) As String
    Dim s As String
    s = UCase$(code)
    If InStr(s, "ACH") > 0 Or InStr(s, "IAT") > 0 Then
        InstrumentOf = "ACH"
    ElseIf InStr(s, "CTRC") > 0 Then
        InstrumentOf = "WIRE"
    ElseIf InStr(s, "WIRE") > 0 Then
        InstrumentOf = "WIRE"
    ElseIf InStr(s, "FED_") > 0 Then
        InstrumentOf = "WIRE"
    ElseIf InStr(s, "SWIFT") > 0 Then
        InstrumentOf = "WIRE"
    End If
End Function

Private Sub AddToList(ByRef s As String, ByVal item As String)
    If item = "" Then Exit Sub
    If InStr(1, "; " & s & ";", "; " & item & ";", vbTextCompare) > 0 Then Exit Sub
    If s = "" Then
        s = item
    Else
        s = s & "; " & item
    End If
End Sub

' =========================================================================
' [PQC] CHECKING AND SHADING THE WRITE-UP
' =========================================================================
Private Function ReviewWriteUp(ByVal srcPath As String, ByVal outDir As String, ByRef fx As PreQCFacts, _
                               ByVal items As Collection) As String
    Dim doc As Object
    Dim outPath As String, fName As String, errDesc As String

    fName = GetFileName(srcPath)
    outPath = outDir & fName
    On Error GoTo Failed
    If FileExists(outPath) Then Kill outPath
    FileCopy srcPath, outPath

    Set doc = GetWord().Documents.Open(FileName:=outPath, ReadOnly:=False, AddToRecentFiles:=False)
    doc.TrackRevisions = False
    CheckWriteUp doc, fx, items
    ProofReviewedCopy doc, fx.Customer & " " & fx.RuleText & " " & fx.Program & " " & fx.SelfCPs & " " & _
                      fx.CustCountry & " " & DictWords(fx.AlertCPs) & " " & DictWords(fx.NonCPs) & " " & _
                      DictWords(fx.CPAmt), items
    InsertPreQCLegend doc
    doc.Save
    doc.Close SaveChanges:=0
    Set doc = Nothing
    ReviewWriteUp = outPath
    Exit Function

Failed:
    errDesc = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=0
    On Error GoTo 0
    AddPQItem items, "Write-up", fName, "", PQV_NA, "could not be reviewed (" & errDesc & ")"
End Function

Private Sub CheckWriteUp(ByVal doc As Object, ByRef fx As PreQCFacts, ByVal items As Collection)
    Dim segs As Collection, seg As Variant
    Dim t As String, segStart As Long
    Dim n As Long, activityDone As Boolean

    Set segs = LineSegments(doc)
    n = 0
    For Each seg In segs
        segStart = CLng(seg(0))
        t = CStr(seg(1))
        n = n + 1
        If n <= 12 Then CheckHeaderLine doc, fx, items, segStart, t
        If InStr(1, t, "alerted due to", vbTextCompare) > 0 Then
            CheckTriggerLine doc, fx, items, segStart, t
        ElseIf IsActivityLine(t, activityDone) Then
            activityDone = True
            CheckActivityLine doc, fx, items, segStart, t
        Else
            SweepTokens doc, fx, items, segStart, t
        End If
    Next seg

    CheckCounterparties doc, fx, items
    CheckOtherActivity fx, items
    If fx.SelfCPs <> "" Then
        AddPQItem items, "Data quality", "", fx.SelfCPs, PQV_EYE, _
                  "counterparty equals the customer's own name on alerted rows - confirm this is an internal transfer"
    End If
End Sub

' "Alert ID:", "Customer:", "Program:", "Rule(s) Triggered:", "Country Risk Rating:"
Private Sub CheckHeaderLine(ByVal doc As Object, ByRef fx As PreQCFacts, ByVal items As Collection, _
                            ByVal segStart As Long, ByVal t As String)
    Dim label As String, val As String, pos As Long, vStart As Long
    Dim ok As Boolean

    pos = InStr(t, ":")
    If pos = 0 Then Exit Sub
    If pos > 30 Then Exit Sub
    label = LCase$(Trim$(Left$(t, pos - 1)))
    val = CleanSegment(Mid$(t, pos + 1))
    If val = "" Then Exit Sub
    vStart = segStart + pos + PadBefore(Mid$(t, pos + 1))

    Select Case label
        Case "alert id"
            ok = False
            If fx.AlertNum <> "" Then ok = (InStr(1, val, fx.AlertNum, vbTextCompare) > 0)
            JudgeHeader doc, items, vStart, val, "Alert ID", fx.AlertNum, ok
        Case "customer"
            JudgeHeader doc, items, vStart, val, "Customer", fx.Customer, SameName(val, fx.Customer)
        Case "rule(s) triggered", "rules triggered", "rule triggered"
            ok = False
            If fx.RuleText <> "" Then ok = (InStr(1, val, fx.RuleText, vbTextCompare) > 0)
            JudgeHeader doc, items, vStart, val, "Rule(s) Triggered", fx.RuleText, ok
        Case "program"
            If fx.Program = "" Then
                AddPQItem items, "Program", val, "", PQV_NA, "no SettlementProgram party in the file"
                ShadeAt doc, vStart, Len(val), PQ_NA
            Else
                If ProgramMatches(val, fx.Program) Then
                    JudgeHeader doc, items, vStart, val, "Program", fx.Program, True
                Else
                    AddPQItem items, "Program", val, fx.Program, PQV_EYE, "the settlement party in the file differs"
                    ShadeAt doc, vStart, Len(val), PQ_EYE
                End If
            End If
        Case "country risk rating"
            If fx.CustCountry = "" Then
                AddPQItem items, "Country", val, "", PQV_NA, "no customer country in the file"
                ShadeAt doc, vStart, Len(val), PQ_NA
            Else
                JudgeHeader doc, items, vStart, val, "Country", fx.CustCountry, CountryMatches(val, fx.CustCountry)
            End If
    End Select
End Sub

Private Sub JudgeHeader(ByVal doc As Object, ByVal items As Collection, ByVal vStart As Long, _
                        ByVal val As String, ByVal label As String, ByVal fileVal As String, ByVal ok As Boolean)
    If fileVal = "" Then
        AddPQItem items, label, val, "", PQV_NA, "the file has no value to compare"
        ShadeAt doc, vStart, Len(val), PQ_NA
    ElseIf ok Then
        AddPQItem items, label, val, fileVal, PQV_OK, ""
        ShadeAt doc, vStart, Len(val), PQ_OK
    Else
        AddPQItem items, label, val, fileVal, PQV_BAD, "does not match the file"
        ShadeAt doc, vStart, Len(val), PQ_BAD
    End If
End Sub

' "... covering the period from D1 to D2"
Private Sub CheckTriggerLine(ByVal doc As Object, ByRef fx As PreQCFacts, ByVal items As Collection, _
                             ByVal segStart As Long, ByVal t As String)
    Dim pos As Long, n As Long, seen As Long
    Dim tok As String, want As String

    pos = 1
    Do
        pos = FindDateToken(t, pos, Len(t), n)
        If pos = 0 Then Exit Do
        seen = seen + 1
        tok = Mid$(t, pos, n)
        If seen = 1 Then
            want = UsDate(fx.AlertMin)
        Else
            want = UsDate(fx.AlertMax)
        End If
        JudgeToken doc, items, segStart + pos - 1, tok, "Review period date " & seen, want
        pos = pos + n
    Loop
End Sub

' The one sentence that states the alerted figures: "... totaling $X ...". Later
' paragraphs also use "totaling" without a figure, so an amount is required and
' only the first such sentence is treated as the activity line.
Private Function IsActivityLine(ByVal t As String, ByVal alreadyDone As Boolean) As Boolean
    Dim pos As Long, n As Long
    If alreadyDone Then Exit Function
    pos = InStr(1, t, "totaling", vbTextCompare)
    If pos = 0 Then Exit Function
    IsActivityLine = (FindAmountToken(t, pos, Len(t), n) > 0)
End Function

' "Between D1 and D2, there were N <direction> <instrument> transfers totaling $X,
'  sent/received by CUSTOMER through their PROGRAM account #ACCT ..."
Private Sub CheckActivityLine(ByVal doc As Object, ByRef fx As PreQCFacts, ByVal items As Collection, _
                              ByVal segStart As Long, ByVal t As String)
    Dim pos As Long, n As Long, seen As Long
    Dim tok As String, want As String
    Dim followWords As Variant
    Dim lt As String, wantDir As String, gotDir As String
    Dim wantIns As String, gotIns As String, gotInsWord As String

    ' Only the wording before "totaling" describes the transactions; the rest of the
    ' paragraph is due-diligence prose that may mention "wire" for other reasons.
    lt = LCase$(Replace(t, Chr(160), " "))
    pos = InStr(lt, "totaling")
    If pos > 0 Then lt = Left$(lt, pos - 1)
    followWords = Array("transaction", "transfer", "txn", "credit", "debit", "payment", "wire", "deposit")

    pos = 1
    Do
        pos = FindDateToken(t, pos, Len(t), n)
        If pos = 0 Then Exit Do
        seen = seen + 1
        tok = Mid$(t, pos, n)
        If seen = 1 Then
            want = UsDate(fx.AlertMin)
        Else
            want = UsDate(fx.AlertMax)
        End If
        JudgeToken doc, items, segStart + pos - 1, tok, "Activity date " & seen, want
        pos = pos + n
    Loop

    pos = FindCountToken(t, 1, Len(t), followWords, n)
    If pos > 0 Then
        tok = Mid$(t, pos, n)
        JudgeToken doc, items, segStart + pos - 1, tok, "Transaction count", UsNumber(fx.AlertCount)
    End If

    pos = AmountAfterTotaling(t, 1, Len(t), n)
    If pos > 0 Then
        tok = Mid$(t, pos, n)
        JudgeToken doc, items, segStart + pos - 1, tok, "Total amount", "$" & UsAmount(fx.AlertTotal)
    End If

    If fx.DrCr = "CR" Then
        wantDir = "incoming"
    ElseIf fx.DrCr = "DR" Then
        wantDir = "outgoing"
    End If
    If InStr(lt, "incoming") > 0 Then gotDir = "incoming"
    If InStr(lt, "outgoing") > 0 Then gotDir = "outgoing"
    If wantDir <> "" Then
        If gotDir = "" Then
            AddPQItem items, "Direction", "", wantDir, PQV_NA, "no direction word in the sentence"
        Else
            JudgeWord doc, items, segStart, t, gotDir, "Direction", wantDir, (gotDir = wantDir)
        End If
    End If

    wantIns = fx.Instrument
    If InStr(lt, " ach ") > 0 Then
        gotIns = "ACH"
        gotInsWord = "ACH"
    End If
    If InStr(lt, "wire") > 0 Then
        gotIns = "WIRE"
        gotInsWord = "wire"
    End If
    If wantIns <> "" Then
        If gotIns = "" Then
            AddPQItem items, "Instrument", "", wantIns, PQV_NA, "no ACH or wire wording in the sentence"
        Else
            JudgeWord doc, items, segStart, t, gotInsWord, "Instrument", wantIns, (gotIns = wantIns)
        End If
    End If

    If fx.AcctPrefix <> "" Then
        pos = InStr(1, t, fx.AcctPrefix, vbTextCompare)
        If pos > 0 Then
            AddPQItem items, "Account number", fx.AcctPrefix, fx.AcctPrefix, PQV_OK, ""
            ShadeAt doc, segStart + pos - 1, Len(fx.AcctPrefix), PQ_OK
        ElseIf InStr(t, "#") > 0 Then
            pos = InStr(t, "#")
            n = AccountTokenLen(t, pos)
            AddPQItem items, "Account number", Mid$(t, pos, n), fx.AcctPrefix, PQV_BAD, _
                      "does not match the account in the file"
            ShadeAt doc, segStart + pos - 1, n, PQ_BAD
        End If
    End If

    If fx.Ccy <> "" Then
        If fx.Ccy <> "USD" Then
            AddPQItem items, "Currency", "$", fx.Ccy, PQV_EYE, "the file is not in USD but the write-up uses $"
        End If
    End If
End Sub

Private Function AccountTokenLen(ByVal t As String, ByVal pos As Long) As Long
    Dim j As Long
    j = pos + 1
    Do While j <= Len(t)
        If Not (Mid$(t, j, 1) Like "[A-Za-z0-9-]") Then Exit Do
        j = j + 1
    Loop
    AccountTokenLen = j - pos
End Function

' Any other amount, date or transaction count anywhere in the document
Private Sub SweepTokens(ByVal doc As Object, ByRef fx As PreQCFacts, ByVal items As Collection, _
                        ByVal segStart As Long, ByVal t As String)
    Dim pos As Long, n As Long, tok As String, lbl As String
    Dim followWords As Variant

    followWords = Array("transaction", "transfer", "txn", "credit", "debit", "payment", "wire", "deposit")

    pos = 1
    Do
        pos = FindAmountToken(t, pos, Len(t), n)
        If pos = 0 Then Exit Do
        tok = Mid$(t, pos, n)
        lbl = BindAmount(fx, tok)
        BindToken doc, items, segStart + pos - 1, tok, "Amount", lbl
        pos = pos + n
    Loop

    pos = 1
    Do
        pos = FindDateToken(t, pos, Len(t), n)
        If pos = 0 Then Exit Do
        tok = Mid$(t, pos, n)
        lbl = BindDate(fx, tok)
        BindToken doc, items, segStart + pos - 1, tok, "Date", lbl
        pos = pos + n
    Loop

    pos = 1
    Do
        pos = FindCountToken(t, pos, Len(t), followWords, n)
        If pos = 0 Then Exit Do
        tok = Mid$(t, pos, n)
        If Not LooksLikeYear(tok) Then
            lbl = BindCount(fx, tok)
            BindToken doc, items, segStart + pos - 1, tok, "Count", lbl
        End If
        pos = pos + n
    Loop
End Sub

' A bare four-digit 19xx/20xx is a year in the prose ("Regulations 2011"), not a count
Private Function LooksLikeYear(ByVal tok As String) As Boolean
    Dim v As Double
    If Len(tok) <> 4 Then Exit Function
    If InStr(tok, ",") > 0 Then Exit Function
    If Not IsNumeric(tok) Then Exit Function
    v = CDbl(tok)
    LooksLikeYear = (v >= 1900 And v <= 2100)
End Function

Private Sub BindToken(ByVal doc As Object, ByVal items As Collection, ByVal absPos As Long, _
                      ByVal tok As String, ByVal kind As String, ByVal label As String)
    If label = "" Then
        AddPQItem items, kind & " (elsewhere)", tok, "", PQV_EYE, "no figure in the file matches this"
        ShadeAt doc, absPos, Len(tok), PQ_EYE
    Else
        AddPQItem items, kind & " (elsewhere)", tok, label, PQV_OK, ""
        ShadeAt doc, absPos, Len(tok), PQ_OK
    End If
End Sub

Private Function BindAmount(ByRef fx As PreQCFacts, ByVal tok As String) As String
    Dim v As Double
    v = ValueToNumber(tok)
    If SameMoney(v, fx.AlertTotal) Then
        BindAmount = "alerted total"
    ElseIf SameMoney(v, fx.AllTotal) Then
        BindAmount = "total of all transactions"
    ElseIf SameMoney(v, fx.NonTotal) Then
        BindAmount = "non-alerted total"
    Else
        BindAmount = CPAmountLabel(fx, v)
    End If
End Function

Private Function CPAmountLabel(ByRef fx As PreQCFacts, ByVal v As Double) As String
    Dim keys As Variant, k As Variant
    If DictCount(fx.CPAmt) = 0 Then Exit Function
    keys = DictKeys(fx.CPAmt)
    For Each k In keys
        If SameMoney(CDbl(DictGet(fx.CPAmt, CStr(k))), v) Then
            CPAmountLabel = "total for " & CPDisplay(fx, CStr(k))
            Exit Function
        End If
    Next k
End Function

Private Function CPDisplay(ByRef fx As PreQCFacts, ByVal key As String) As String
    If DictExists(fx.AlertCPs, key) Then
        CPDisplay = CStr(DictGet(fx.AlertCPs, key))
    Else
        CPDisplay = key
    End If
End Function

Private Function BindDate(ByRef fx As PreQCFacts, ByVal tok As String) As String
    If tok = UsDate(fx.AlertMin) Then
        BindDate = "first alerted transaction"
    ElseIf tok = UsDate(fx.AlertMax) Then
        BindDate = "last alerted transaction"
    End If
End Function

Private Function BindCount(ByRef fx As PreQCFacts, ByVal tok As String) As String
    Dim v As Double
    v = ValueToNumber(tok)
    If v = fx.AlertCount Then
        BindCount = "alerted transaction count"
    ElseIf v = DictCount(fx.AlertCPs) Then
        BindCount = "alerted counterparty count"
    ElseIf v = fx.NonCount Then
        BindCount = "non-alerted transaction count"
    ElseIf v = fx.AllCount Then
        BindCount = "count of all transactions"
    End If
End Function

Private Sub JudgeToken(ByVal doc As Object, ByVal items As Collection, ByVal absPos As Long, _
                       ByVal tok As String, ByVal label As String, ByVal want As String)
    If want = "" Then
        AddPQItem items, label, tok, "", PQV_NA, "the file has no value to compare"
        ShadeAt doc, absPos, Len(tok), PQ_NA
    ElseIf NormValue(tok) = NormValue(want) Then
        AddPQItem items, label, tok, want, PQV_OK, ""
        ShadeAt doc, absPos, Len(tok), PQ_OK
    Else
        AddPQItem items, label, tok, want, PQV_BAD, "the file says " & want
        ShadeAt doc, absPos, Len(tok), PQ_BAD
    End If
End Sub

Private Sub JudgeWord(ByVal doc As Object, ByVal items As Collection, ByVal segStart As Long, ByVal t As String, _
                      ByVal word As String, ByVal label As String, ByVal want As String, ByVal ok As Boolean)
    Dim pos As Long
    pos = InStr(1, t, word, vbTextCompare)
    If ok Then
        AddPQItem items, label, word, want, PQV_OK, ""
        If pos > 0 Then ShadeAt doc, segStart + pos - 1, Len(word), PQ_OK
    Else
        AddPQItem items, label, word, want, PQV_BAD, "the file says " & want
        If pos > 0 Then ShadeAt doc, segStart + pos - 1, Len(word), PQ_BAD
    End If
End Sub

' Every alerted counterparty: named in the write-up or not
Private Sub CheckCounterparties(ByVal doc As Object, ByRef fx As PreQCFacts, ByVal items As Collection)
    Dim docKey As String
    Dim keys As Variant, k As Variant
    Dim cp As String, share As Double, amt As Double
    Dim named As Long, total As Long

    total = DictCount(fx.AlertCPs)
    If total = 0 Then Exit Sub
    docKey = NormCP(doc.Content.Text)

    keys = DictKeys(fx.AlertCPs)
    For Each k In keys
        cp = CStr(DictGet(fx.AlertCPs, CStr(k)))
        amt = 0
        If DictExists(fx.CPAmt, CStr(k)) Then amt = CDbl(DictGet(fx.CPAmt, CStr(k)))
        share = 0
        If fx.AlertTotal <> 0 Then share = amt / fx.AlertTotal
        If InStr(1, docKey, CStr(k), vbTextCompare) > 0 Then
            named = named + 1
            AddPQItem items, "Counterparty named", cp, "$" & UsAmount(amt), PQV_OK, ""
            ShadeFirst doc, cp, PQ_OK
        ElseIf share >= PQ_CP_SHARE Then
            AddPQItem items, "Counterparty missing", "", cp & " ($" & UsAmount(amt) & ")", PQV_BAD, _
                      Format$(share * 100, "0") & "% of the alerted total and not named in the write-up"
        Else
            AddPQItem items, "Counterparty not named", "", cp & " ($" & UsAmount(amt) & ")", PQV_NA, _
                      "below " & Format$(PQ_CP_SHARE * 100, "0") & "% of the alerted total"
        End If
    Next k
    If named = 0 Then
        AddPQItem items, "Counterparty coverage", "0 of " & UsNumber(total), "", PQV_EYE, _
                  "no alerted counterparty is named in the write-up"
    End If
End Sub

' Supports the "same counterparties ... plus N additional" wording in OTHER ACTIVITY REVIEW
Private Sub CheckOtherActivity(ByRef fx As PreQCFacts, ByVal items As Collection)
    Dim keys As Variant, k As Variant
    Dim inBoth As Long, extra As Long, verdict As String, tail As String

    If DictCount(fx.NonCPs) = 0 Then Exit Sub
    keys = DictKeys(fx.NonCPs)
    For Each k In keys
        If DictExists(fx.AlertCPs, CStr(k)) Then
            inBoth = inBoth + 1
        Else
            extra = extra + 1
        End If
    Next k
    If inBoth > 0 Then
        verdict = PQV_OK
    Else
        verdict = PQV_EYE
    End If
    If extra = 1 Then
        tail = "1 additional counterparty"
    Else
        tail = UsNumber(extra) & " additional counterparties"
    End If
    AddPQItem items, "Other activity counterparties", "", _
              UsNumber(inBoth) & " also in the alert period, " & tail, verdict, _
              "the Other Activity Review should describe " & tail
End Sub

' =========================================================================
' [PQC] SMALL COMPARISONS
' =========================================================================
Private Function SameMoney(ByVal a As Double, ByVal b As Double) As Boolean
    SameMoney = (Abs(a - b) < 0.005)
End Function

Private Function SameName(ByVal a As String, ByVal b As String) As Boolean
    Dim ka As String, kb As String
    ka = NormCP(a)
    kb = NormCP(b)
    If ka = "" Then Exit Function
    If kb = "" Then Exit Function
    If ka = kb Then
        SameName = True
    ElseIf InStr(1, ka, kb, vbTextCompare) > 0 Then
        SameName = True
    ElseIf InStr(1, kb, ka, vbTextCompare) > 0 Then
        SameName = True
    End If
End Function

' "Airwallex" against "AIRWALLEX US LLC ORIG", "Currencycloud" against "Currency Cloud Stlmnt".
' fileVal may list several parties separated by "; ".
Private Function ProgramMatches(ByVal narrVal As String, ByVal fileVal As String) As Boolean
    Dim parts() As String, i As Long
    parts = Split(fileVal, "; ")
    For i = 0 To UBound(parts)
        If OneProgramMatches(narrVal, parts(i)) Then
            ProgramMatches = True
            Exit Function
        End If
    Next i
End Function

Private Function OneProgramMatches(ByVal narrVal As String, ByVal fileVal As String) As Boolean
    Dim a As String, b As String
    a = Replace(NormCP(narrVal), " ", "")
    b = Replace(NormCP(fileVal), " ", "")
    If a = "" Then Exit Function
    If b = "" Then Exit Function
    If InStr(1, b, a, vbTextCompare) > 0 Then
        OneProgramMatches = True
    ElseIf InStr(1, a, b, vbTextCompare) > 0 Then
        OneProgramMatches = True
    ElseIf Len(a) >= 5 Then
        OneProgramMatches = (InStr(1, b, Left$(a, 5), vbTextCompare) > 0)
    End If
End Function

' "Low (UK)" against a file country of "GB"
Private Function CountryMatches(ByVal narrVal As String, ByVal fileVal As String) As Boolean
    Dim inside As String, p1 As Long, p2 As Long
    Dim a As String, b As String

    p1 = InStr(narrVal, "(")
    p2 = InStr(narrVal, ")")
    inside = narrVal
    If p1 > 0 Then
        If p2 > p1 Then inside = Mid$(narrVal, p1 + 1, p2 - p1 - 1)
    End If
    a = CountryKey(inside)
    b = CountryKey(fileVal)
    If a = "" Then Exit Function
    If b = "" Then Exit Function
    CountryMatches = (a = b)
End Function

Private Function CountryKey(ByVal s As String) As String
    Dim k As String
    k = UCase$(Trim$(s))
    Select Case k
        Case "UK", "GB", "GBR", "UNITED KINGDOM"
            CountryKey = "GB"
        Case "USA", "US", "UNITED STATES"
            CountryKey = "US"
        Case Else
            CountryKey = k
    End Select
End Function

' Header value without control characters or trailing punctuation
Private Function CleanSegment(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = Chr(160) Then ch = " "
        If AscW(ch) >= 32 Then out = out & ch
    Next i
    out = Trim$(out)
    Do While Len(out) > 0
        If Not (Right$(out, 1) = "." Or Right$(out, 1) = ",") Then Exit Do
        out = Trim$(Left$(out, Len(out) - 1))
    Loop
    CleanSegment = out
End Function

Private Function PadBefore(ByVal s As String) As Long
    Dim i As Long, n As Long
    n = 0
    For i = 1 To Len(s)
        If Mid$(s, i, 1) <> " " And Mid$(s, i, 1) <> Chr(160) Then Exit For
        n = n + 1
    Next i
    PadBefore = n
End Function

' =========================================================================
' [PQC] WORD SHADING
' =========================================================================
' Shades a token whose position was worked out as paragraph start + index in the paragraph text
Private Sub ShadeAt(ByVal doc As Object, ByVal absPos As Long, ByVal tokLen As Long, ByVal colr As Long)
    If tokLen <= 0 Then Exit Sub
    ShadeExact doc, TruePos(doc, absPos, tokLen), tokLen, colr
End Sub

' Shades an exact document range (positions that came from Word's Find)
Private Sub ShadeExact(ByVal doc As Object, ByVal absPos As Long, ByVal tokLen As Long, ByVal colr As Long)
    Dim rng As Object
    If tokLen <= 0 Then Exit Sub
    On Error Resume Next
    Set rng = doc.Range(absPos, absPos + tokLen)
    If Not rng Is Nothing Then rng.Shading.BackgroundPatternColor = colr
    On Error GoTo 0
End Sub

' Word counts the hidden code of fields - and hyperlinks are fields - in document positions,
' but Range.Text leaves it out. A position worked out as paragraph start + index in the
' paragraph's text therefore lands too early after a hyperlink. When the paragraph has such
' hidden characters, the token is found again with Find: the same occurrence of the same text,
' counted from the start of the paragraph. Paragraphs without them are returned unchanged.
Private Function TruePos(ByVal doc As Object, ByVal absPos As Long, ByVal tokLen As Long) As Long
    Dim pr As Object, pText As String, k As Long, tok As String, occ As Long, p As Long
    Dim hits As Collection

    TruePos = absPos
    On Error GoTo Done
    Set pr = doc.Range(absPos, absPos).Paragraphs(1).Range
    pText = pr.Text
    If Len(pText) = pr.End - pr.Start Then Exit Function
    k = absPos - pr.Start
    If k < 0 Or k + tokLen > Len(pText) Then Exit Function
    tok = Mid$(pText, k + 1, tokLen)
    ' occurrence number of the token at index k, counted the way Find counts (case-insensitive,
    ' non-overlapping)
    occ = 0
    p = InStr(1, pText, tok, vbTextCompare)
    Do While p > 0
        occ = occ + 1
        If p >= k + 1 Then Exit Do
        p = InStr(p + Len(tok), pText, tok, vbTextCompare)
    Loop
    If p <> k + 1 Then Exit Function
    Set hits = WdFindIn(doc, pr.Start, pr.End, tok, False)
    If hits.Count >= occ Then TruePos = CLng(hits(occ)(0))
Done:
End Function

' Shades the first mention of findText that has no shading yet (a Financial Institutions
' or Subject line may already carry its own verdict colour)
Private Sub ShadeFirst(ByVal doc As Object, ByVal findText As String, ByVal colr As Long)
    Dim hits As Collection, h As Variant, cur As Long
    If findText = "" Then Exit Sub
    Set hits = WdFindIn(doc, doc.Content.Start, doc.Content.End, findText, False)
    For Each h In hits
        cur = -16777216
        On Error Resume Next
        cur = doc.Range(CLng(h(0)), CLng(h(1))).Shading.BackgroundPatternColor
        On Error GoTo 0
        If cur = -16777216 Or cur = 16777215 Then       ' wdColorAutomatic / white
            ShadeExact doc, CLng(h(0)), CLng(h(1)) - CLng(h(0)), colr
            Exit For
        End If
    Next h
End Sub

Private Sub InsertPreQCLegend(ByVal doc As Object)
    Dim rng As Object, para As Object
    Dim txt As String

    txt = "Pre-QC colour key - green: matches the transaction file. Red: contradicts the file. " & _
          "Yellow: needs a reviewer's eye. Grey: examined but not verifiable from the file. " & _
          "Unshaded text was not checked. This copy is for review only."
    If PQ_PROOF Then
        txt = txt & " Spelling and grammar: wavy red underline = possible misspelling, wavy blue underline = " & _
              "Word grammar flag; words and spacing the tool corrected are listed on the findings sheet " & _
              "(corrected words shaded yellow)."
    End If
    If InStr(1, doc.Paragraphs(1).Range.Text, "Pre-QC colour key", vbTextCompare) > 0 Then Exit Sub

    doc.Paragraphs(1).Range.InsertParagraphBefore
    Set para = doc.Paragraphs(1)
    para.Range.InsertBefore txt

    Set para = doc.Paragraphs(1)
    Set rng = doc.Range(para.Range.Start, para.Range.End - 1)
    On Error Resume Next
    rng.Style = "Normal"
    On Error GoTo 0
    With rng.Font
        .Bold = False
        .Italic = True
        .Underline = 0                       ' wdUnderlineNone
        .Color = 0                           ' black
        .Size = 9
    End With
    rng.ParagraphFormat.Alignment = 0        ' wdAlignParagraphLeft
    rng.Shading.BackgroundPatternColor = PQ_EYE
End Sub

' =========================================================================
' [PQC] REPORT SHEETS
' =========================================================================
Private Sub AddPQItem(ByVal items As Collection, ByVal label As String, ByVal narrVal As String, _
                      ByVal fileVal As String, ByVal verdict As String, ByVal note As String)
    items.Add Array(label, narrVal, fileVal, verdict, note)
End Sub

Private Sub CountVerdicts(ByVal items As Collection, ByRef nOK As Long, ByRef nBad As Long, ByRef nEye As Long)
    Dim it As Variant
    nOK = 0
    nBad = 0
    nEye = 0
    For Each it In items
        Select Case CStr(it(3))
            Case PQV_OK
                nOK = nOK + 1
            Case PQV_BAD
                nBad = nBad + 1
            Case PQV_EYE
                nEye = nEye + 1
        End Select
    Next it
End Sub

Private Sub PreQCPrepareSheets(ByVal wsDash As Worksheet, ByVal wsFind As Worksheet, ByVal title As String, _
                               ByVal dashHeaders As Variant)
    wsDash.Cells.Clear
    wsFind.Cells.Clear
    wsDash.Range("B2").Value = title
    wsFind.Range("B2").Value = title & " - FINDINGS"
    StylePQTitle wsDash.Range("B2")
    StylePQTitle wsFind.Range("B2")
    WriteHeaderRow wsDash, 3, 2, dashHeaders
    WriteHeaderRow wsFind, 3, 2, Array("ECM ID", "Alert ID", "Check", "Narrative Says", "File Says", _
                                       "Verdict", "Note")
End Sub

Private Sub StylePQTitle(ByVal c As Range)
    With c
        .Font.Name = FONT_UI
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = COLOR_TEXT_DARK
    End With
End Sub

Private Sub WritePreQCDash(ByVal ws As Worksheet, ByVal rowIdx As Long, ByRef fx As PreQCFacts, _
                           ByVal nOK As Long, ByVal nBad As Long, ByVal nEye As Long, _
                           ByVal docPath As String, ByVal outPath As String)
    Dim status As String

    If docPath = "" Then
        status = "No write-up found"
    ElseIf Not fx.HasData Then
        status = "No alerted rows in the file"
    ElseIf nBad > 0 Then
        status = "Mismatch - review"
    ElseIf nEye > 0 Then
        status = "Check flagged items"
    Else
        status = "Clean"
    End If

    ws.Cells(rowIdx, 2).Value = fx.EcmID
    ws.Cells(rowIdx, 3).Value = fx.AlertID
    ws.Cells(rowIdx, 4).Value = fx.Customer
    ws.Cells(rowIdx, 5).Value = fx.RuleText
    ws.Cells(rowIdx, 6).Value = fx.AlertCount
    ws.Cells(rowIdx, 7).Value = fx.AlertTotal
    ws.Cells(rowIdx, 8).Value = DateText(fx.AlertMin) & " - " & DateText(fx.AlertMax)
    ws.Cells(rowIdx, 9).Value = nOK
    ws.Cells(rowIdx, 10).Value = nBad
    ws.Cells(rowIdx, 11).Value = nEye
    ws.Cells(rowIdx, 12).Value = status
    ws.Cells(rowIdx, 13).Value = GetFileName(docPath)
    ws.Cells(rowIdx, 14).Value = GetFileName(outPath)
    ws.Cells(rowIdx, 7).NumberFormat = "$#,##0.00"

    If nBad > 0 Then
        HighlightCell ws.Cells(rowIdx, 12), COLOR_ALERT_RED, COLOR_FILL_RED
        HighlightCell ws.Cells(rowIdx, 10), COLOR_ALERT_RED, COLOR_FILL_RED
    ElseIf nEye > 0 Then
        HighlightCell ws.Cells(rowIdx, 12), COLOR_AMBER, COLOR_FILL_AMBER
    Else
        HighlightCell ws.Cells(rowIdx, 12), COLOR_SUCCESS_GREEN, COLOR_FILL_GREEN
    End If
End Sub

Private Sub WritePreQCFindings(ByVal ws As Worksheet, ByRef rowIdx As Long, ByVal ecmID As String, _
                               ByVal alertID As String, ByVal items As Collection)
    Dim it As Variant, verdict As String
    For Each it In items
        verdict = CStr(it(3))
        If verdict <> PQV_OK Then
            ws.Cells(rowIdx, 2).Value = ecmID
            ws.Cells(rowIdx, 3).Value = alertID
            ws.Cells(rowIdx, 4).Value = CStr(it(0))
            ws.Cells(rowIdx, 5).Value = CStr(it(1))
            ws.Cells(rowIdx, 6).Value = CStr(it(2))
            ws.Cells(rowIdx, 7).Value = verdict
            ws.Cells(rowIdx, 8).Value = CStr(it(4))
            Select Case verdict
                Case PQV_BAD
                    HighlightCell ws.Cells(rowIdx, 7), COLOR_ALERT_RED, COLOR_FILL_RED
                Case PQV_EYE
                    HighlightCell ws.Cells(rowIdx, 7), COLOR_AMBER, COLOR_FILL_AMBER
                Case PQV_PROOF
                    HighlightCell ws.Cells(rowIdx, 7), COLOR_ACCENT_BLUE, COLOR_CARD_BG
                Case Else
                    HighlightCell ws.Cells(rowIdx, 7), COLOR_TEXT_MUTED, COLOR_CARD_BG
            End Select
            rowIdx = rowIdx + 1
        End If
    Next it
End Sub

Private Sub FinishPreQCSheet(ByVal ws As Worksheet, ByVal nCols As Long, ByVal lastRow As Long, _
                             ByVal emptyMsg As String)
    Dim c As Long, endRow As Long

    endRow = lastRow
    If endRow < 4 Then
        ws.Cells(4, 2).Value = emptyMsg
        ws.Cells(4, 2).Font.Italic = True
        ws.Cells(4, 2).Font.Color = COLOR_TEXT_MUTED
        endRow = 4
    End If
    With ws.Range(ws.Cells(4, 2), ws.Cells(endRow, nCols + 1))
        .Font.Name = FONT_UI
        .Font.Size = 9.5
        .VerticalAlignment = xlTop
    End With
    For c = 2 To nCols + 1
        ws.Columns(c).AutoFit
        If ws.Columns(c).ColumnWidth > 46 Then ws.Columns(c).ColumnWidth = 46
    Next c
    SetSheetZoom ws, SHEET_ZOOM
End Sub

' First amount after the word "totaling" between fromPos and toPos, or the first
' amount at all when the word is absent. "62 transfers for amounts ranging from
' $305.00 to $99,998.00 totaling $1,775,116.00" must give $1,775,116.00.
Private Function AmountAfterTotaling(ByVal t As String, ByVal fromPos As Long, ByVal toPos As Long, _
                                     ByRef tokLen As Long) As Long
    Dim tp As Long, p As Long
    If fromPos < 1 Then fromPos = 1
    If toPos > Len(t) Then toPos = Len(t)
    If toPos < fromPos Then Exit Function
    tp = InStr(fromPos, t, "totaling", vbTextCompare)
    If tp > 0 And tp < toPos Then
        p = FindAmountToken(t, tp, toPos, tokLen)
        If p > 0 Then
            AmountAfterTotaling = p
            Exit Function
        End If
    End If
    AmountAfterTotaling = FindAmountToken(t, fromPos, toPos, tokLen)
End Function

' =========================================================================
' [PQC] PRE-QC FOR ESCALATION NARRATIVES
' =========================================================================
' Reads each lookback transaction file (plus the alerted transaction file when
' one sits next to it) with its escalation narrative and checks every figure,
' sentence by sentence:
'   - opening sentence and Total Suspicious Dollar Amount: overall activity,
'     whole dollars rounded up
'   - Date Range of Suspicious Activity and the "Between ..." line: the period in
'     the lookback file name (NARRATIVE_DATE_RANGE = "FILE")
'   - "Between ..." first figures: all, credit or debit activity, with cents
'   - every other "N transactions totaling $X" claim is recomputed from the rows,
'     using the date window and the counterparty named in that sentence
'   - ACH vs wire wording against the transaction codes of the rows described
'   - Subject, program partner, account number and banks against the file
'   - customer name spelling, own-name transfers, the alerted activity and the
'     "all high dollar" / "consecutive days" wording
' The headline facts come from LoadTxnFile, so pre-QC and the updater read the
' same numbers.
' =========================================================================
Public Sub Run_PreQC_Escalation()
    Dim st As AppState
    Dim lookDir As String, narrDir As String, outDir As String, note As String
    Dim jobs As Collection, job As Variant
    Dim wsDash As Worksheet, wsFind As Worksheet
    Dim items As Collection
    Dim dashRow As Long, findRow As Long
    Dim nOK As Long, nBad As Long, nEye As Long
    Dim okTot As Long, badTot As Long, eyeTot As Long, nDocs As Long
    Dim nFix As Long, nRev As Long, fixTot As Long, revTot As Long
    Dim outPath As String, msg As String, docPath As String, alertInfo As String
    Dim ctx As NarrCtx, blankCtx As NarrCtx
    Dim wbOut As Workbook

    Set wbOut = PreQCBook()
    If wbOut Is Nothing Then Exit Sub
    lookDir = Pick_Folder("Select the folder with the lookback transaction files")
    If lookDir = "" Then Exit Sub

    Set jobs = PreQCEscalationJobs(lookDir)
    If jobs.Count = 0 Then
        MsgBox "No lookback transaction file (ECMID_ALERTID_...xlsx) was found in that folder.", vbExclamation, "Pre-QC"
        Exit Sub
    End If

    ' Narratives next to the lookback files, or in a folder of their own
    narrDir = ""
    For Each job In jobs
        If FindNarrativeFile(lookDir, CStr(job(1)), CStr(job(2)), note) <> "" Then
            narrDir = lookDir
            Exit For
        End If
    Next job
    If narrDir = "" Then
        narrDir = Pick_Folder("Select the folder with the escalation narratives")
        If narrDir = "" Then Exit Sub
    End If
    outDir = PreQCOutDir(narrDir)

    SaveAndSpeedUp st
    On Error GoTo Failed

    Set wsDash = GetOrCreateWorksheet(wbOut, SHEET_PQ_ESC)
    Set wsFind = GetOrCreateWorksheet(wbOut, SHEET_PQ_ESC_FIND)
    PreQCPrepareSheets wsDash, wsFind, "PRE-QC: ESCALATION NARRATIVES", _
                       Array("ECM ID", "Alert ID", "Lookback File", "Transactions", "Total Amount", _
                             "Narrative Period", "Alerted", "Verified", "Mismatch", "To Check", "Status", _
                             "Narrative", "Reviewed Copy", "Spelling / Grammar")
    dashRow = 4
    findRow = 4

    For Each job In jobs
        nDocs = nDocs + 1
        Set items = New Collection
        ctx = blankCtx
        alertInfo = ""
        docPath = FindNarrativeFile(narrDir, CStr(job(1)), CStr(job(2)), note)
        outPath = RunEscalationJob(job, docPath, outDir, ctx, items, alertInfo)
        If note <> "" Then AddPQItem items, "Narrative", GetFileName(docPath), "", PQV_EYE, note
        CountVerdicts items, nOK, nBad, nEye
        okTot = okTot + nOK
        badTot = badTot + nBad
        eyeTot = eyeTot + nEye
        WriteEscDash wsDash, dashRow, ctx, CStr(job(1)), CStr(job(2)), alertInfo, nOK, nBad, nEye, docPath, outPath
        wsDash.Cells(dashRow, 15).Value = ProofSummary(items)
        ProofCounts items, nFix, nRev
        fixTot = fixTot + nFix
        revTot = revTot + nRev
        WritePreQCFindings wsFind, findRow, CStr(job(1)), CStr(job(2)), items
        dashRow = dashRow + 1
    Next job

    FinishPreQCSheet wsDash, 14, dashRow - 1, "No alerts were checked."
    FinishPreQCSheet wsFind, 7, findRow - 1, "Nothing flagged - every checked figure matched the file."
    ReleaseWord
    RestoreApp st
    wsDash.Activate
    PreQCDoneMessage nDocs, okTot, badTot, eyeTot, outDir, "Escalation", fixTot, revTot
    Exit Sub

Failed:
    msg = Err.Description
    ReleaseWord
    RestoreApp st
    MsgBox "Pre-QC stopped: " & msg, vbCritical, "Pre-QC"
End Sub

' Lookback files in the folder: names containing "Lookback", or every Excel file with
' ECM/alert IDs when none does. Combined closure files are skipped. Each job is
' Array(lookback path, ECM ID, alert ID, alerted-file path or "").
Private Function PreQCEscalationJobs(ByVal folderPath As String) As Collection
    Dim jobs As New Collection, anyLookback As Boolean
    Dim files As Collection, f As Variant, alerted As Object
    Dim fName As String, ecmID As String, alertID As String, matchKey As String, key As String, aPath As String

    Set files = ListExcelFiles(folderPath, "")
    Set alerted = CreateLookupDict()
    For Each f In files
        fName = GetFileName(CStr(f))
        If InStr(1, fName, "lookback", vbTextCompare) > 0 Then
            anyLookback = True
        ElseIf IsAlertedFileName(fName) Then
            ExtractFileIDs CStr(f), ecmID, alertID, matchKey
            If ecmID <> "" Then
                key = NarrativePrefix(ecmID, alertID)
                If Not DictExists(alerted, key) Then DictAdd alerted, key, CStr(f)
            End If
        End If
    Next f
    For Each f In files
        fName = GetFileName(CStr(f))
        If InStr(1, fName, "Combined", vbTextCompare) = 0 And Not IsAlertedFileName(fName) Then
            If (Not anyLookback) Or InStr(1, fName, "lookback", vbTextCompare) > 0 Then
                ExtractFileIDs CStr(f), ecmID, alertID, matchKey
                If ecmID <> "" Then
                    key = NarrativePrefix(ecmID, alertID)
                    aPath = ""
                    If DictExists(alerted, key) Then aPath = CStr(DictGet(alerted, key))
                    jobs.Add Array(CStr(f), ecmID, alertID, aPath)
                End If
            End If
        End If
    Next f
    Set PreQCEscalationJobs = jobs
End Function

' "..._Alerted_Transaction.xlsx": an alerted-transactions extract, not a lookback file
Private Function IsAlertedFileName(ByVal fName As String) As Boolean
    Dim s As String
    s = LCase$(fName)
    If InStr(s, "alerted") = 0 Then Exit Function
    If InStr(s, "lookback") > 0 Or InStr(s, "combined") > 0 Or InStr(s, "non alerted") > 0 Or InStr(s, "non-alerted") > 0 Then Exit Function
    IsAlertedFileName = True
End Function

Private Function RunEscalationJob(ByVal job As Variant, ByVal docPath As String, ByVal outDir As String, _
                                  ByRef ctx As NarrCtx, ByVal items As Collection, ByRef alertInfo As String) As String
    Dim ts As TxnSet, ats As TxnSet, pr As PairResult, ef As EscFacts, aef As EscFacts
    Dim errDesc As String, customer As String

    On Error GoTo Failed
    LoadTxnFile CStr(job(0)), ts, CreateLookupDict()
    BuildNarrCtx ctx, CStr(job(1)), CStr(job(2)), CStr(job(0)), CStr(job(0)), ts, ts, pr
    LoadEscFacts CStr(job(0)), ts, ef
    InitEscFacts aef
    If CStr(job(3)) <> "" Then
        LoadTxnFile CStr(job(3)), ats, CreateLookupDict()
        AttachAlertedFile ts, ats, ef, GetFileName(CStr(job(3)))
        LoadEscRows CStr(job(3)), aef
        alertInfo = UsNumber(ef.AlertN) & " / $" & UsAmount(ef.AlertTotal) & " (alerted file)"
    ElseIf ts.AlertCount > 0 Then
        alertInfo = UsNumber(ts.AlertCount) & " / $" & UsAmount(ts.AlertAmt) & " (lookback flag)"
    End If
    If ts.UniqueCount = 0 Then
        AddPQItem items, "Lookback file", GetFileName(CStr(job(0))), "", PQV_EYE, "no transactions were read from the file"
    End If
    If docPath = "" Then
        AddPQItem items, "Narrative", "", "", PQV_NA, "no narrative found for this alert"
    Else
        customer = CustomerFromFileName(docPath)
        RunEscalationJob = ReviewEscalation(docPath, outDir, ctx, ef, aef, items, customer)
    End If
    Exit Function
Failed:
    errDesc = Err.Description
    AddPQItem items, "Lookback file", GetFileName(CStr(job(0))), "", PQV_EYE, "could not be read (" & errDesc & ")"
End Function

Private Sub AttachAlertedFile(ByRef ts As TxnSet, ByRef ats As TxnSet, ByRef ef As EscFacts, ByVal fName As String)
    Dim keys As Variant, k As Variant
    ef.HasAlertFile = True
    ef.AlertFile = fName
    ef.AlertN = ats.UniqueCount
    ef.AlertTotal = ats.TotalAmt
    ef.AlertMin = ats.MinDate
    ef.AlertMax = ats.MaxDate
    ef.AlertRule = ats.RuleText
    ef.LookbackAlerted = ts.AlertCount
    ef.AlertMissing = 0
    If DictCount(ats.Map) > 0 Then
        keys = DictKeys(ats.Map)
        For Each k In keys
            If Not DictExists(ts.Map, CStr(k)) Then ef.AlertMissing = ef.AlertMissing + 1
        Next k
    End If
    AddAmtFact ef, ats.TotalAmt, "alerted file total"
    AddAmtFact ef, -Int(-Abs(ats.TotalAmt)), "alerted file total rounded up"
    AddCntFact ef, ats.UniqueCount, "alerted file count"
End Sub

' =========================================================================
' [PQC] ESCALATION FACT BOOK
' =========================================================================
Private Sub InitEscFacts(ByRef ef As EscFacts)
    Set ef.AmtSet = CreateLookupDict()
    Set ef.CntSet = CreateLookupDict()
    Set ef.DateSet = CreateLookupDict()
    Set ef.Groups = CreateLookupDict()
    Set ef.CPs = CreateLookupDict()
    Set ef.CPAmt = CreateLookupDict()
    Set ef.Banks = CreateLookupDict()
    Set ef.Holders = CreateLookupDict()
    Set ef.Families = CreateLookupDict()
    ef.TxN = 0
End Sub

' Rows of the alerted transaction file, for the counterparty selection
Private Sub LoadEscRows(ByVal fPath As String, ByRef ef As EscFacts)
    Dim wb As Workbook, ws As Worksheet
    Dim openedByUs As Boolean
    InitEscFacts ef
    Set wb = GetOrOpenWorkbook(fPath, openedByUs)
    On Error GoTo Cleanup
    Set ws = FindTransactionSheet(wb)
    If Not ws Is Nothing Then ReadEscSheet ws, ef
Cleanup:
    On Error Resume Next
    If openedByUs Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub LoadEscFacts(ByVal fPath As String, ByRef ts As TxnSet, ByRef ef As EscFacts)
    Dim wb As Workbook, ws As Worksheet
    Dim openedByUs As Boolean

    InitEscFacts ef
    ef.Total = ts.TotalAmt

    Set wb = GetOrOpenWorkbook(fPath, openedByUs)
    On Error GoTo Cleanup
    Set ws = FindTransactionSheet(wb)
    If Not ws Is Nothing Then ReadEscSheet ws, ef

    ' Headline figures, identical to what the narrative updater writes
    AddAmtFact ef, ts.TotalAmt, "total of all transactions"
    AddAmtFact ef, -Int(-Abs(ts.TotalAmt)), "total rounded up"
    AddAmtFact ef, ts.CrAmt, "credit total"
    AddAmtFact ef, ts.DrAmt, "debit total"
    AddAmtFact ef, ts.AlertAmt, "alerted total"
    AddAmtFact ef, ts.MinAmt, "smallest transaction"
    AddAmtFact ef, ts.MaxAmt, "largest transaction"
    If ts.UniqueCount > 0 Then AddAmtFact ef, Round(ts.TotalAmt / ts.UniqueCount, 2), "average transaction"
    If ts.CrCount > 0 Then AddAmtFact ef, Round(ts.CrAmt / ts.CrCount, 2), "average credit"
    If ts.DrCount > 0 Then AddAmtFact ef, Round(ts.DrAmt / ts.DrCount, 2), "average debit"
    AddCntFact ef, ts.UniqueCount, "count of all transactions"
    AddCntFact ef, ts.CrCount, "credit count"
    AddCntFact ef, ts.DrCount, "debit count"
    AddCntFact ef, ts.AlertCount, "alerted count"
    AddCntFact ef, DictCount(ef.CPs), "distinct counterparties"
    AddGroupFacts ef
    FinishEscFacts ef

Cleanup:
    On Error Resume Next
    If openedByUs Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub ReadEscSheet(ByVal ws As Worksheet, ByRef ef As EscFacts)
    Dim block As Variant, seen As Object
    Dim lastRow As Long, lastCol As Long, hdrRow As Long, cap As Long
    Dim cTx As Long, cAmt As Long, cTxDate As Long, cCP As Long, cDrCr As Long, cCode As Long, cDesc As Long
    Dim cAcct As Long, cOrig As Long, cBen As Long, cAl As Long, cRule As Long
    Dim cPT(1 To 5) As Long, cPN(1 To 5) As Long
    Dim r As Long, q As Long, okDate As Boolean
    Dim tx As String, amt As Double, dv As Date, cp As String, cpKey As String, dr As String, dirKey As String
    Dim code As String, pType As String, pName As String, holder As String, ins As String

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedCol(ws)
    If lastRow < 2 Then Exit Sub
    If lastCol < 1 Then Exit Sub
    hdrRow = FindHeaderRow(ws, FLD_TXID)
    If hdrRow = 0 Then Exit Sub
    If hdrRow >= lastRow Then Exit Sub

    block = ReadBlock(ws, hdrRow, 1, lastRow, lastCol)
    cTx = NamedCol(block, lastCol, "Transaction ID")
    cAmt = NamedCol(block, lastCol, "Transaction Amount")
    cTxDate = NamedCol(block, lastCol, "Transaction Date")
    cCP = NamedCol(block, lastCol, "Counterparty")
    cDrCr = NamedCol(block, lastCol, "Dr Cr")
    cCode = NamedCol(block, lastCol, "Transaction Code")
    cDesc = NamedCol(block, lastCol, "Transaction Code Description")
    cAcct = NamedCol(block, lastCol, "Account No")
    cOrig = NamedCol(block, lastCol, "Originator Name")
    cBen = NamedCol(block, lastCol, "Beneficiary Name")
    cAl = NamedCol(block, lastCol, "Is Alerted Transaction?")
    If cAl = 0 Then cAl = NamedCol(block, lastCol, "Is Alerted Transaction")
    cRule = NamedCol(block, lastCol, "Alert Information")
    For q = 1 To 5
        cPT(q) = NamedCol(block, lastCol, "Party Type" & q)
        cPN(q) = NamedCol(block, lastCol, "Party Name" & q)
    Next q
    If cTx = 0 Or cAmt = 0 Then Exit Sub

    cap = UBound(block, 1)
    ReDim ef.TxDate(1 To cap)
    ReDim ef.TxHasDate(1 To cap)
    ReDim ef.TxAmt(1 To cap)
    ReDim ef.TxCP(1 To cap)
    ReDim ef.TxDir(1 To cap)
    ReDim ef.TxIns(1 To cap)
    ReDim ef.TxAl(1 To cap)
    ReDim ef.TxRule(1 To cap)

    Set seen = CreateLookupDict()
    For r = 2 To UBound(block, 1)
        tx = Trim$(CellText(BlockCell(block, r, cTx)))
        If tx <> "" Then
            If Not DictExists(seen, tx) Then
                DictAdd seen, tx, True
                amt = Abs(ValueToNumber(BlockCell(block, r, cAmt)))
                dv = ParseDateValue(BlockCell(block, r, cTxDate), okDate)
                dr = UCase$(Trim$(CellText(BlockCell(block, r, cDrCr))))
                dirKey = DeriveCounterparty(dr, "C", "D")
                cp = Trim$(CellText(BlockCell(block, r, cCP)))
                If cCP = 0 Then cp = DeriveCounterparty(dr, CellText(BlockCell(block, r, cOrig)), CellText(BlockCell(block, r, cBen)))
                cpKey = NormCP(cp)
                code = Trim$(CellText(BlockCell(block, r, cCode)))
                ins = InstrumentOf(code)
                If ins = "" Then ins = InstrumentOf(CellText(BlockCell(block, r, cDesc)))

                ef.TxN = ef.TxN + 1
                ef.TxAmt(ef.TxN) = amt
                ef.TxHasDate(ef.TxN) = okDate
                If okDate Then ef.TxDate(ef.TxN) = dv
                ef.TxCP(ef.TxN) = cpKey
                ef.TxDir(ef.TxN) = dirKey
                ef.TxIns(ef.TxN) = ins
                ef.TxAl(ef.TxN) = IsYes(CellText(BlockCell(block, r, cAl)))
                ef.TxRule(ef.TxN) = Trim$(CellText(BlockCell(block, r, cRule)))

                AddAmtFact ef, amt, "a single transaction"
                If okDate Then
                    If Not DictExists(ef.DateSet, Format$(dv, "yyyymmdd")) Then DictAdd ef.DateSet, Format$(dv, "yyyymmdd"), "transaction date"
                End If
                If ef.AcctPrefix = "" Then
                    ef.AcctPrefix = Trim$(CellText(BlockCell(block, r, cAcct)))
                    If InStr(ef.AcctPrefix, "-") > 0 Then ef.AcctPrefix = Left$(ef.AcctPrefix, InStr(ef.AcctPrefix, "-") - 1)
                End If

                ' Account holder: the beneficiary of a credit, the originator of a debit
                holder = NormCP(DeriveCounterparty(dr, CellText(BlockCell(block, r, cBen)), CellText(BlockCell(block, r, cOrig))))
                If holder <> "" Then AddToAmount ef.Holders, holder, 1

                If cpKey <> "" Then
                    If Not DictExists(ef.CPs, cpKey) Then DictAdd ef.CPs, cpKey, cp
                    AddToAmount ef.CPAmt, cpKey, amt
                    AddGroup ef, "counterparty " & cpKey, amt
                    If dirKey <> "" Then AddGroup ef, "counterparty " & cpKey & " (" & dirKey & ")", amt
                End If
                If code <> "" Then
                    AddGroup ef, "code " & code, amt
                    If dirKey <> "" Then AddGroup ef, "code " & code & " (" & dirKey & ")", amt
                End If
                If okDate Then AddGroup ef, "month " & Format$(dv, "yyyy-mm"), amt
                If amt > 0 Then
                    If Abs(amt - 100 * Int(amt / 100 + 0.0000001)) < 0.005 Then AddGroup ef, "round-dollar (100s)", amt
                    If Abs(amt - 1000 * Int(amt / 1000 + 0.0000001)) < 0.005 Then AddGroup ef, "round-dollar (1000s)", amt
                End If

                For q = 1 To 5
                    pType = LCase$(Trim$(CellText(BlockCell(block, r, cPT(q)))))
                    pName = Trim$(CellText(BlockCell(block, r, cPN(q))))
                    If pType <> "" And pName <> "" Then
                        If pType = "settlementprogram" Then
                            AddToList ef.Programs, pName
                        Else
                            AddToAmount ef.Banks, NormCP(pName), amt
                        End If
                    End If
                Next q
            End If
        End If
    Next r
End Sub

' Most frequent account holder, and first words used by 2+ counterparties
' ("AMAZON" for AMAZON.C1BWFDWAB, AMAZON MEXICO SERVICES INC, ...)
Private Sub FinishEscFacts(ByRef ef As EscFacts)
    Dim keys As Variant, k As Variant, best As Double, w As String, firstWords As Object

    If DictCount(ef.Holders) > 0 Then
        keys = DictKeys(ef.Holders)
        For Each k In keys
            If CDbl(DictGet(ef.Holders, CStr(k))) > best Then
                best = CDbl(DictGet(ef.Holders, CStr(k)))
                ef.Holder = CStr(k)
            End If
        Next k
    End If

    Set firstWords = CreateLookupDict()
    If DictCount(ef.CPs) > 0 Then
        keys = DictKeys(ef.CPs)
        For Each k In keys
            w = FirstWord(CStr(k))
            If Len(w) >= 4 Then AddToAmount firstWords, w, 1
        Next k
        keys = DictKeys(firstWords)
        For Each k In keys
            If CDbl(DictGet(firstWords, CStr(k))) >= 2 Then DictAdd ef.Families, CStr(k), True
        Next k
    End If
End Sub

Private Function FirstWord(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStr(s, " ")
    If p > 0 Then FirstWord = Left$(s, p - 1) Else FirstWord = s
End Function

Private Sub AddGroup(ByRef ef As EscFacts, ByVal label As String, ByVal amt As Double)
    Dim v As Variant
    If DictExists(ef.Groups, label) Then
        v = DictGet(ef.Groups, label)
        DictSet ef.Groups, label, Array(CLng(v(0)) + 1, CDbl(v(1)) + amt)
    Else
        DictAdd ef.Groups, label, Array(CLng(1), amt)
    End If
End Sub

Private Sub AddToAmount(ByVal d As Object, ByVal key As String, ByVal amt As Double)
    If key = "" Then Exit Sub
    If DictExists(d, key) Then
        DictSet d, key, CDbl(DictGet(d, key)) + amt
    Else
        DictAdd d, key, amt
    End If
End Sub

Private Sub AddGroupFacts(ByRef ef As EscFacts)
    Dim keys As Variant, k As Variant, v As Variant
    If DictCount(ef.Groups) = 0 Then Exit Sub
    keys = DictKeys(ef.Groups)
    For Each k In keys
        v = DictGet(ef.Groups, CStr(k))
        AddCntFact ef, CLng(v(0)), CStr(k) & " count"
        AddAmtFact ef, CDbl(v(1)), CStr(k) & " total"
    Next k
End Sub

Private Function CentsKey(ByVal x As Double) As String
    CentsKey = Format$(Round(Abs(x) * 100, 0), "0")
End Function

Private Sub AddAmtFact(ByRef ef As EscFacts, ByVal x As Double, ByVal label As String)
    Dim k As String
    k = CentsKey(x)
    If Not DictExists(ef.AmtSet, k) Then DictAdd ef.AmtSet, k, label
End Sub

Private Sub AddCntFact(ByRef ef As EscFacts, ByVal n As Long, ByVal label As String)
    Dim k As String
    k = CStr(n)
    If Not DictExists(ef.CntSet, k) Then DictAdd ef.CntSet, k, label
End Sub

Private Function AmtFactLabel(ByRef ef As EscFacts, ByVal v As Double) As String
    Dim k As String
    k = CentsKey(v)
    If DictExists(ef.AmtSet, k) Then AmtFactLabel = CStr(DictGet(ef.AmtSet, k))
End Function

Private Function CntFactLabel(ByRef ef As EscFacts, ByVal n As Long) As String
    Dim k As String
    k = CStr(n)
    If DictExists(ef.CntSet, k) Then CntFactLabel = CStr(DictGet(ef.CntSet, k))
End Function

' =========================================================================
' [PQC] ESCALATION: SUBSETS OF TRANSACTIONS
' =========================================================================
' A claim's scope: an optional date window, an optional counterparty ("=KEY" for one
' counterparty, "^WORD" for every counterparty starting with WORD) and an optional
' direction ("C"/"D").
Private Sub SubsetStats(ByRef ef As EscFacts, ByVal hasWin As Boolean, ByVal wFrom As Date, ByVal wTo As Date, _
                        ByVal cpOpt As String, ByVal dirOpt As String, _
                        ByRef n As Long, ByRef amt As Double, ByRef nAch As Long, ByRef nWire As Long)
    Dim i As Long, k As String, fam As String
    n = 0
    amt = 0
    nAch = 0
    nWire = 0
    If cpOpt <> "" Then
        k = Mid$(cpOpt, 2)
        fam = k & " "
    End If
    For i = 1 To ef.TxN
        If TxInScope(ef, i, hasWin, wFrom, wTo, cpOpt, k, fam, dirOpt) Then
            n = n + 1
            amt = amt + ef.TxAmt(i)
            If ef.TxIns(i) = "ACH" Then nAch = nAch + 1
            If ef.TxIns(i) = "WIRE" Then nWire = nWire + 1
        End If
    Next i
End Sub

Private Function TxInScope(ByRef ef As EscFacts, ByVal i As Long, ByVal hasWin As Boolean, ByVal wFrom As Date, _
                           ByVal wTo As Date, ByVal cpOpt As String, ByVal k As String, ByVal fam As String, _
                           ByVal dirOpt As String) As Boolean
    If hasWin Then
        If Not ef.TxHasDate(i) Then Exit Function
        If ef.TxDate(i) < wFrom Or ef.TxDate(i) > wTo Then Exit Function
    End If
    If cpOpt <> "" Then
        If Left$(cpOpt, 1) = "=" Then
            If ef.TxCP(i) <> k Then Exit Function
        Else
            If ef.TxCP(i) <> k And Left$(ef.TxCP(i), Len(fam)) <> fam Then Exit Function
        End If
    End If
    If dirOpt <> "" Then
        If ef.TxDir(i) <> dirOpt Then Exit Function
    End If
    TxInScope = True
End Function

Private Function ScopeText(ByVal hasWin As Boolean, ByVal wFrom As Date, ByVal wTo As Date, _
                           ByVal cpOpt As String, ByVal dirOpt As String) As String
    Dim s As String
    If dirOpt = "C" Then s = "credits"
    If dirOpt = "D" Then s = "debits"
    If cpOpt <> "" Then
        If s <> "" Then s = s & ", "
        If Left$(cpOpt, 1) = "=" Then s = s & Mid$(cpOpt, 2) Else s = s & Mid$(cpOpt, 2) & "*"
    End If
    If hasWin Then
        If s <> "" Then s = s & ", "
        s = s & UsDate(wFrom) & "-" & UsDate(wTo)
    End If
    If s = "" Then s = "all activity"
    ScopeText = s
End Function

' Counterparties named in a sentence: exact names, then name families ("Amazon")
Private Function CPOptions(ByRef ef As EscFacts, ByVal s As String) As Collection
    Dim col As New Collection, key As String, keys As Variant, k As Variant, famHit As Object

    Set CPOptions = col
    key = " " & NormCP(s) & " "
    Set famHit = CreateLookupDict()
    If DictCount(ef.CPs) > 0 Then
        keys = DictKeys(ef.CPs)
        For Each k In keys
            If CStr(k) <> "" Then
                If InStr(key, " " & CStr(k) & " ") > 0 Then
                    col.Add "=" & CStr(k)
                    If Not DictExists(famHit, FirstWord(CStr(k))) Then DictAdd famHit, FirstWord(CStr(k)), True
                End If
            End If
        Next k
    End If
    If DictCount(ef.Families) > 0 Then
        keys = DictKeys(ef.Families)
        For Each k In keys
            If InStr(key, " " & CStr(k) & " ") > 0 And Not DictExists(famHit, CStr(k)) Then col.Add "^" & CStr(k)
        Next k
    End If
End Function

' =========================================================================
' [PQC] ESCALATION: TEXT HELPERS
' =========================================================================
' End of the sentence starting at fromPos (the position of its full stop, or Len(t)).
' "Inc.", "Ltd.", "N.A." and single initials do not end a sentence.
Private Function PQSentenceEnd(ByVal t As String, ByVal fromPos As Long) As Long
    Dim i As Long, j As Long, nextCh As String, w As String
    For i = fromPos To Len(t)
        If Mid$(t, i, 1) = "." Then
            nextCh = Mid$(t, i + 1, 1)
            If nextCh = " " Or nextCh = "" Or nextCh = vbCr Or nextCh = vbLf Or nextCh = Chr(160) Then
                j = i - 1
                Do While j >= fromPos
                    If Not (Mid$(t, j, 1) Like "[A-Za-z.]") Then Exit Do
                    j = j - 1
                Loop
                w = LCase$(Mid$(t, j + 1, i - j - 1))
                Do While Right$(w, 1) = "."
                    w = Left$(w, Len(w) - 1)
                Loop
                If Not IsAbbreviation(w) Then
                    PQSentenceEnd = i
                    Exit Function
                End If
            End If
        End If
    Next i
    PQSentenceEnd = Len(t)
End Function

Private Function IsAbbreviation(ByVal w As String) As Boolean
    If Len(w) = 1 Then
        IsAbbreviation = True
        Exit Function
    End If
    Select Case w
        Case "inc", "ltd", "llc", "co", "corp", "mr", "ms", "mrs", "dr", "st", "no", "jr", "sr", _
             "u.s", "n.a", "e.g", "i.e", "al", "vs"
            IsAbbreviation = True
    End Select
End Function

Private Function NumberWordValue(ByVal w As String) As Long
    Dim words As Variant, i As Long
    words = Array("one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", _
                  "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", _
                  "nineteen", "twenty")
    For i = 0 To UBound(words)
        If w = CStr(words(i)) Then
            NumberWordValue = i + 1
            Exit Function
        End If
    Next i
End Function

' True when one of the next three words after position j starts with a follow word
Private Function FollowedBy(ByVal t As String, ByVal j As Long, ByVal followWords As Variant) As Boolean
    Dim rest As String, words() As String, w As Variant, fw As Variant, n As Long
    rest = LCase$(Replace(Mid$(t, j, 60), Chr(160), " "))
    rest = Replace(Replace(Replace(Replace(Replace(Replace(rest, vbCr, " "), ",", " "), ".", " "), ";", " "), ":", " "), "(", " ")
    If Trim$(rest) = "" Then Exit Function
    words = Split(Application.WorksheetFunction.Trim(rest), " ")
    For Each w In words
        If CStr(w) <> "" Then
            n = n + 1
            For Each fw In followWords
                If Left$(CStr(w), Len(CStr(fw))) = CStr(fw) Then
                    FollowedBy = True
                    Exit Function
                End If
            Next fw
            If n >= 3 Then Exit For
        End If
    Next w
End Function

' First count written as a word ("five wire transactions"); returns its position (0 = none)
Private Function FindCountWord(ByVal t As String, ByVal fromPos As Long, ByVal toPos As Long, _
                               ByVal followWords As Variant, ByRef tokLen As Long, ByRef value As Long) As Long
    Dim i As Long, j As Long, prevCh As String, w As String, v As Long
    i = fromPos
    Do While i <= toPos
        If Mid$(t, i, 1) Like "[A-Za-z]" Then
            If i > 1 Then prevCh = Mid$(t, i - 1, 1) Else prevCh = " "
            j = i
            Do While j <= Len(t)
                If Not (Mid$(t, j, 1) Like "[A-Za-z]") Then Exit Do
                j = j + 1
            Loop
            If Not (prevCh Like "[A-Za-z]") Then
                w = LCase$(Mid$(t, i, j - i))
                v = NumberWordValue(w)
                If v > 0 Then
                    If FollowedBy(t, j, followWords) Then
                        tokLen = j - i
                        value = v
                        FindCountWord = i
                        Exit Function
                    End If
                End If
            End If
            i = j
        Else
            i = i + 1
        End If
    Loop
End Function

' Next count in digits or words, whichever comes first
Private Function NextCount(ByVal t As String, ByVal fromPos As Long, ByVal toPos As Long, _
                           ByVal followWords As Variant, ByRef tokLen As Long, ByRef value As Long) As Long
    Dim p1 As Long, l1 As Long, p2 As Long, l2 As Long, v2 As Long
    p1 = FindCountToken(t, fromPos, toPos, followWords, l1)
    p2 = FindCountWord(t, fromPos, toPos, followWords, l2, v2)
    If p1 > 0 And (p2 = 0 Or p1 < p2) Then
        tokLen = l1
        value = CLng(ValueToNumber(Mid$(t, p1, l1)))
        NextCount = p1
    ElseIf p2 > 0 Then
        tokLen = l2
        value = v2
        NextCount = p2
    End If
End Function

' "wire"/"ACH" within the two words after a count; returns the claimed instrument
Private Function InstrumentClaim(ByVal t As String, ByVal afterPos As Long, ByRef wordPos As Long, _
                                 ByRef wordLen As Long) As String
    Dim rest As String, words() As String, i As Long, n As Long, w As String, claim As String
    rest = Replace(Mid$(t, afterPos, 40), Chr(160), " ")
    rest = Replace(Replace(Replace(Replace(rest, ",", " "), ".", " "), ";", " "), ":", " ")
    If Trim$(rest) = "" Then Exit Function
    words = Split(Application.WorksheetFunction.Trim(rest), " ")
    For i = 0 To UBound(words)
        w = LCase$(words(i))
        If w <> "" Then
            n = n + 1
            If Left$(w, 4) = "wire" Then
                claim = "WIRE"
            ElseIf w = "ach" Or w = "iat" Then
                claim = "ACH"
            End If
            If claim <> "" Then
                wordPos = InStr(afterPos, t, words(i), vbTextCompare)
                wordLen = Len(words(i))
                InstrumentClaim = claim
                Exit Function
            End If
            If n >= 2 Then Exit For
        End If
    Next i
End Function

' Date tokens between lo and hi as Array(position, length, date)
Private Function SentenceDates(ByVal t As String, ByVal lo As Long, ByVal hi As Long) As Collection
    Dim col As New Collection, pos As Long, n As Long, d As Date, ok As Boolean
    pos = lo
    Do
        pos = FindDateToken(t, pos, hi, n)
        If pos = 0 Then Exit Do
        d = ParseDateValue(Mid$(t, pos, n), ok)
        If ok Then col.Add Array(pos, n, d)
        pos = pos + n
    Loop
    Set SentenceDates = col
End Function

' Date window for a count at countPos: the last complete date pair before it, else the
' first pair after it, else "on <date>" before it
Private Function ClaimWindow(ByVal t As String, ByVal dts As Collection, ByVal countPos As Long, _
                             ByRef wFrom As Date, ByRef wTo As Date) As Boolean
    Dim i As Long, found As Boolean, d As Variant
    For i = 1 To dts.Count - 1 Step 2
        If CLng(dts(i + 1)(0)) + CLng(dts(i + 1)(1)) <= countPos Then
            wFrom = dts(i)(2)
            wTo = dts(i + 1)(2)
            found = True
        End If
    Next i
    If Not found And dts.Count >= 2 Then
        wFrom = dts(1)(2)
        wTo = dts(2)(2)
        found = True
    End If
    If Not found Then
        For Each d In dts
            If CLng(d(0)) < countPos And CLng(d(0)) > 3 Then
                If LCase$(Mid$(t, CLng(d(0)) - 3, 3)) = "on " Then
                    wFrom = d(2)
                    wTo = d(2)
                    found = True
                End If
            End If
        Next d
    End If
    ClaimWindow = found
End Function

' =========================================================================
' [PQC] ESCALATION: NAMES
' =========================================================================
Private Function Distinctive(ByVal key As String) As Collection
    Dim col As New Collection, parts() As String, i As Long, w As String
    Set Distinctive = col
    If Trim$(key) = "" Then Exit Function
    parts = Split(Trim$(key), " ")
    For i = 0 To UBound(parts)
        w = parts(i)
        If Len(w) >= 3 And Not IsLegalSuffix(w) And Not (w Like "*#*") Then col.Add w
    Next i
End Function

Private Function IsLegalSuffix(ByVal w As String) As Boolean
    Select Case w
        Case "LTD", "LLC", "INC", "LIMITED", "COMPANY", "CORP", "CORPORATION", "CO", "PTY", "GMBH", "LLP", _
             "PLC", "THE", "AND"
            IsLegalSuffix = True
    End Select
End Function

' Same person or company written differently: "DUSTIN HALIMAN" / "DUSTIN PATRICK HALIMAN",
' "EVELYN NAVARRO MORALES" / "NAVARRO MORALES EVELYN"
Private Function SameParty(ByVal a As String, ByVal b As String) As Boolean
    Dim wa As Collection, wb As Collection, x As Variant, y As Variant, nShared As Long
    Set wa = Distinctive(NormCP(a))
    Set wb = Distinctive(NormCP(b))
    If wa.Count = 0 Or wb.Count = 0 Then Exit Function
    For Each x In wa
        For Each y In wb
            If CStr(x) = CStr(y) Then
                nShared = nShared + 1
                Exit For
            End If
        Next y
    Next x
    If nShared >= 2 Then
        SameParty = True
    ElseIf wa.Count = 1 And wb.Count = 1 Then
        SameParty = (wa(1) = wb(1))
    Else
        SameParty = (wa(1) = wb(1) And wa(wa.Count) = wb(wb.Count))
    End If
End Function

' Counterparty named in the narrative: its full name, its first and last distinctive words,
' or its first two ("WELLS FARGO IFI" is named by "Wells Fargo")
Private Function NamedIn(ByVal key As String, ByVal docWords As Object, ByVal docKey As String) As Boolean
    Dim d As Collection
    If InStr(" " & docKey & " ", " " & key & " ") > 0 Then
        NamedIn = True
        Exit Function
    End If
    Set d = Distinctive(key)
    If d.Count = 0 Then Exit Function
    If DictExists(docWords, CStr(d(1))) And DictExists(docWords, CStr(d(d.Count))) Then
        NamedIn = True
    ElseIf d.Count >= 2 Then
        NamedIn = (DictExists(docWords, CStr(d(1))) And DictExists(docWords, CStr(d(2))))
    End If
End Function

Private Function WordSet(ByVal key As String) As Object
    Dim d As Object, parts() As String, i As Long
    Set d = CreateLookupDict()
    If Trim$(key) <> "" Then
        parts = Split(key, " ")
        For i = 0 To UBound(parts)
            If parts(i) <> "" Then
                If Not DictExists(d, parts(i)) Then DictAdd d, parts(i), True
            End If
        Next i
    End If
    Set WordSet = d
End Function

Private Function Levenshtein(ByVal a As String, ByVal b As String) As Long
    Dim la As Long, lb As Long, i As Long, j As Long, cost As Long
    Dim prevRow() As Long, curRow() As Long
    la = Len(a)
    lb = Len(b)
    If Abs(la - lb) > 2 Then
        Levenshtein = 9
        Exit Function
    End If
    ReDim prevRow(0 To lb)
    ReDim curRow(0 To lb)
    For j = 0 To lb
        prevRow(j) = j
    Next j
    For i = 1 To la
        curRow(0) = i
        For j = 1 To lb
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then cost = 0 Else cost = 1
            curRow(j) = prevRow(j) + 1
            If curRow(j - 1) + 1 < curRow(j) Then curRow(j) = curRow(j - 1) + 1
            If prevRow(j - 1) + cost < curRow(j) Then curRow(j) = prevRow(j - 1) + cost
        Next j
        For j = 0 To lb
            prevRow(j) = curRow(j)
        Next j
    Next i
    Levenshtein = prevRow(lb)
End Function

' =========================================================================
' [PQC] CHECKING THE ESCALATION NARRATIVE
' =========================================================================
Private Function ReviewEscalation(ByVal srcPath As String, ByVal outDir As String, ByRef ctx As NarrCtx, _
                                  ByRef ef As EscFacts, ByRef aef As EscFacts, ByVal items As Collection, _
                                  ByVal customer As String) As String
    Dim doc As Object
    Dim outPath As String, fName As String, errDesc As String

    fName = GetFileName(srcPath)
    outPath = outDir & fName
    On Error GoTo Failed
    If FileExists(outPath) Then Kill outPath
    FileCopy srcPath, outPath

    Set doc = GetWord().Documents.Open(FileName:=outPath, ReadOnly:=False, AddToRecentFiles:=False)
    doc.TrackRevisions = False
    CheckEscalation doc, ctx, ef, aef, items, customer
    ProofReviewedCopy doc, customer & " " & ef.Programs & " " & ef.Holder & " " & DictWords(ef.CPs) & " " & _
                      DictWords(ef.Banks) & " " & DictWords(ef.Holders) & " " & DictWords(aef.CPs), items
    InsertPreQCLegend doc
    doc.Save
    doc.Close SaveChanges:=0
    Set doc = Nothing
    ReviewEscalation = outPath
    Exit Function

Failed:
    errDesc = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=0
    On Error GoTo 0
    AddPQItem items, "Narrative", fName, "", PQV_EYE, "could not be reviewed (" & errDesc & ")"
End Function

Private Sub CheckEscalation(ByVal doc As Object, ByRef ctx As NarrCtx, ByRef ef As EscFacts, ByRef aef As EscFacts, _
                            ByVal items As Collection, ByVal customer As String)
    Dim segs As Collection, seg As Variant
    Dim t As String, lt As String, head As String, segStart As Long
    Dim escDone As Boolean, totDone As Boolean, drDone As Boolean, progDone As Boolean, subjDone As Boolean
    Dim fiDone As Boolean, betDone As Boolean
    Dim se As Long
    Dim inhFrom As Date, inhTo As Date, hasOwn As Boolean, ownFrom As Date, ownTo As Date

    ef.OnceFlags = ""
    Set segs = LineSegments(doc)
    For Each seg In segs
        segStart = CLng(seg(0))
        t = CStr(seg(1))
        lt = LCase$(LTrim$(t))
        head = Left$(lt, 45)
        If Not escDone And InStr(lt, "escalated") > 0 And InStr(lt, "totaling") > 0 Then
            escDone = True
            se = EscOpening(doc, ctx, items, segStart, t)
            EscSentences doc, ctx, ef, items, segStart, t, se + 1, False, inhFrom, inhTo
        ElseIf Not totDone And InStr(lt, "total suspicious dollar amount") > 0 Then
            totDone = True
            EscTotalLine doc, ctx, items, segStart, t
        ElseIf Not drDone And InStr(lt, "date range of suspicious activity") > 0 Then
            drDone = True
            EscDateRange doc, ctx, items, segStart, t
        ElseIf Not subjDone And InStr(head, "subject") > 0 And InStr(head, ":") > 0 Then
            subjDone = True
            EscSubjectLine doc, ef, items, segStart, t
        ElseIf Not progDone And InStr(lt, "program partner") > 0 Then
            progDone = True
            EscProgramLine doc, ef, items, segStart, t
        ElseIf Not fiDone And InStr(head, "financial institution") > 0 And InStr(Left$(t, 60), ":") > 0 Then
            fiDone = True
            EscBankLine doc, ef, items, segStart, t
        ElseIf Not betDone And Left$(lt, 8) = "between " And InStr(lt, "totaling") > 0 Then
            betDone = True
            se = PQSentenceEnd(t, 1)
            hasOwn = EscBetween(doc, ctx, ef, items, segStart, t, se, ownFrom, ownTo)
            EscSentences doc, ctx, ef, items, segStart, t, se + 1, hasOwn, ownFrom, ownTo
        Else
            EscSentences doc, ctx, ef, items, segStart, t, 1, False, inhFrom, inhTo
        End If
    Next seg

    If Not escDone Then AddPQItem items, "Opening sentence", "", "", PQV_EYE, "line not found"
    If Not totDone Then AddPQItem items, "Total Suspicious Dollar Amount", "", "", PQV_EYE, "line not found"
    If Not drDone Then AddPQItem items, "Date Range of Suspicious Activity", "", "", PQV_EYE, "line not found"
    EscCounterparties doc, ef, items
    EscSelection doc, ef, aef, items
    EscAlerted doc, ef, items
    EscNameCheck doc, ef, items, customer
End Sub

' Every sentence of t from fromPos on. The first sentence with a date pair becomes the
' window that later sentences without dates ("Out of these transactions ...") inherit.
Private Sub EscSentences(ByVal doc As Object, ByRef ctx As NarrCtx, ByRef ef As EscFacts, ByVal items As Collection, _
                         ByVal segStart As Long, ByVal t As String, ByVal fromPos As Long, _
                         ByVal hasInh As Boolean, ByVal inhFrom As Date, ByVal inhTo As Date)
    Dim s0 As Long, s1 As Long, hasOwn As Boolean, ownFrom As Date, ownTo As Date
    s0 = fromPos
    Do While s0 <= Len(t)
        s1 = PQSentenceEnd(t, s0)
        hasOwn = EscSentence(doc, ctx, ef, items, segStart, t, s0, s1, hasInh, inhFrom, inhTo, "Sentence", ownFrom, ownTo)
        If hasOwn And Not hasInh Then
            hasInh = True
            inhFrom = ownFrom
            inhTo = ownTo
        End If
        s0 = s1 + 1
    Loop
End Sub

' "... to report 104 transactions totaling $7,394,666 conducted by ..."; returns the
' end of the first sentence
Private Function EscOpening(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal items As Collection, _
                            ByVal segStart As Long, ByVal t As String) As Long
    Dim sentEnd As Long, p As Long, n As Long, v As Long, aPos As Long

    sentEnd = PQSentenceEnd(t, 1)
    EscOpening = sentEnd
    p = NextCount(t, 1, sentEnd, Array("transaction", "transfer", "txn"), n, v)
    If p > 0 Then
        If v = ctx.NewCount Then
            MarkToken doc, items, segStart + p - 1, Mid$(t, p, n), "Opening sentence count", UsNumber(ctx.NewCount), PQV_OK, ""
        Else
            MarkToken doc, items, segStart + p - 1, Mid$(t, p, n), "Opening sentence count", UsNumber(ctx.NewCount), _
                      PQV_BAD, "the file says " & UsNumber(ctx.NewCount)
        End If
    End If
    aPos = InStr(1, t, "totaling", vbTextCompare)
    If aPos < 1 Then aPos = 1
    p = FindAmountToken(t, aPos, sentEnd, n)
    If p > 0 Then JudgeWhole doc, items, segStart + p - 1, Mid$(t, p, n), "Opening sentence amount", ctx.NewTotal
End Function

Private Sub EscTotalLine(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal items As Collection, _
                         ByVal segStart As Long, ByVal t As String)
    Dim p As Long, n As Long
    p = FindAmountToken(t, 1, Len(t), n)
    If p > 0 Then JudgeWhole doc, items, segStart + p - 1, Mid$(t, p, n), "Total Suspicious Dollar Amount", ctx.NewTotal
End Sub

Private Sub EscDateRange(ByVal doc As Object, ByRef ctx As NarrCtx, ByVal items As Collection, _
                         ByVal segStart As Long, ByVal t As String)
    Dim p1 As Long, l1 As Long, p2 As Long, l2 As Long
    p1 = FindDateToken(t, 1, Len(t), l1)
    If p1 = 0 Then Exit Sub
    JudgeRuleDate doc, items, segStart + p1 - 1, Mid$(t, p1, l1), "Date range start", ctx.NarrFrom, ctx
    p2 = FindDateToken(t, p1 + l1, Len(t), l2)
    If p2 > 0 Then JudgeRuleDate doc, items, segStart + p2 - 1, Mid$(t, p2, l2), "Date range end", ctx.NarrTo, ctx
End Sub

' "Subject: Midagency LTD" - the account holder in the file should be one of the subjects
Private Sub EscSubjectLine(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, _
                           ByVal segStart As Long, ByVal t As String)
    Dim colon As Long, val As String, vStart As Long, parts() As String, i As Long, ok As Boolean

    colon = InStr(t, ":")
    If colon = 0 Then Exit Sub
    val = CleanSegment(Mid$(t, colon + 1))
    If val = "" Then Exit Sub
    vStart = segStart + colon + PadBefore(Mid$(t, colon + 1))
    parts = Split(Replace(Replace(val, ";", ","), " and ", ","), ",")
    For i = 0 To UBound(parts)
        If Trim$(parts(i)) <> "" And ef.Holder <> "" Then
            If SameParty(parts(i), ef.Holder) Then ok = True
        End If
    Next i
    If ok Then
        AddPQItem items, "Subject", val, ef.Holder, PQV_OK, ""
        ShadeAt doc, vStart, Len(val), PQ_OK
    ElseIf ef.Holder = "" Or InStr(1, val, "unknown", vbTextCompare) > 0 Then
        AddPQItem items, "Subject", val, ef.Holder, PQV_NA, "subject is unknown or the file has no account holder"
        ShadeAt doc, vStart, Len(val), PQ_NA
    Else
        AddPQItem items, "Subject", val, ef.Holder, PQV_EYE, "the account holder in the file is " & ef.Holder
        ShadeAt doc, vStart, Len(val), PQ_EYE
    End If
End Sub

' "5. Associated Program Partner: Currencycloud, #8335028587" or
' "Associated Program Partner: TransferWise INC 2, ..., Account Number: 8314367894"
Private Sub EscProgramLine(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, _
                           ByVal segStart As Long, ByVal t As String)
    Dim colon As Long, vStart As Long, cut As Long, aPos As Long, aLen As Long
    Dim val As String, progName As String, acct As String, digitsN As String, digitsF As String

    colon = InStr(t, ":")
    If colon = 0 Then Exit Sub
    val = Mid$(t, colon + 1)
    vStart = colon + 1 + PadBefore(val)
    progName = Trim$(val)
    cut = InStr(progName, ",")
    If InStr(progName, "#") > 0 Then
        If cut = 0 Or InStr(progName, "#") < cut Then cut = InStr(progName, "#")
    End If
    If cut > 0 Then progName = Trim$(Left$(progName, cut - 1))

    If progName <> "" Then
        If ef.Programs = "" Then
            AddPQItem items, "Program partner", progName, "", PQV_NA, "no SettlementProgram party in the file"
            ShadeAt doc, segStart + vStart - 1, Len(progName), PQ_NA
        ElseIf ProgramMatches(progName, ef.Programs) Then
            AddPQItem items, "Program partner", progName, ef.Programs, PQV_OK, ""
            ShadeAt doc, segStart + vStart - 1, Len(progName), PQ_OK
        Else
            AddPQItem items, "Program partner", progName, ef.Programs, PQV_EYE, "the settlement party in the file differs"
            ShadeAt doc, segStart + vStart - 1, Len(progName), PQ_EYE
        End If
    End If

    aPos = AccountTokenAfterMarker(t, colon, aLen)
    If aPos > 0 And ef.AcctPrefix <> "" Then
        acct = Mid$(t, aPos, aLen)
        digitsN = DigitsOnly(acct)
        digitsF = DigitsOnly(ef.AcctPrefix)
        If digitsN <> "" And InStr(digitsF, digitsN) > 0 Then
            AddPQItem items, "Account number", acct, ef.AcctPrefix, PQV_OK, ""
            ShadeAt doc, segStart + aPos - 1, aLen, PQ_OK
        Else
            AddPQItem items, "Account number", acct, ef.AcctPrefix, PQV_BAD, "does not match the account in the file"
            ShadeAt doc, segStart + aPos - 1, aLen, PQ_BAD
        End If
    End If
End Sub

' Account token after "#", "Account Number", "Account No" or "Acct" (0 = none)
Private Function AccountTokenAfterMarker(ByVal t As String, ByVal fromPos As Long, ByRef tokLen As Long) As Long
    Dim markers As Variant, m As Variant, p As Long, best As Long, bestLen As Long, j As Long, k As Long
    markers = Array("#", "account number", "account no", "acct")
    For Each m In markers
        p = InStr(fromPos, t, CStr(m), vbTextCompare)
        If p > 0 And (best = 0 Or p < best) Then
            best = p
            bestLen = Len(CStr(m))
        End If
    Next m
    If best = 0 Then Exit Function
    j = best + bestLen
    Do While j <= Len(t)
        If Not (Mid$(t, j, 1) Like "[ :#.]") Then Exit Do
        j = j + 1
    Loop
    k = j
    Do While k <= Len(t)
        If Not (Mid$(t, k, 1) Like "[A-Za-z0-9-]") Then Exit Do
        k = k + 1
    Loop
    If k = j Then Exit Function
    If DigitsOnly(Mid$(t, j, k - j)) = "" Then Exit Function
    tokLen = k - j
    AccountTokenAfterMarker = j
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

' "Financial Institutions Involved: CFSB and WELLS FARGO BANK" - every bank that carries
' at least PQ_CP_SHARE of the volume should be named
Private Sub EscBankLine(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, _
                        ByVal segStart As Long, ByVal t As String)
    Dim colon As Long, vStart As Long, val As String, lineKey As String
    Dim keys As Variant, k As Variant, bank As String, missing As String, bigList As String
    Dim share As Double

    colon = InStr(t, ":")
    If colon = 0 Then Exit Sub
    val = CleanSegment(Mid$(t, colon + 1))
    If val = "" Then Exit Sub
    vStart = segStart + colon + PadBefore(Mid$(t, colon + 1))
    If DictCount(ef.Banks) = 0 Or ef.Total = 0 Then
        AddPQItem items, "Financial institutions", val, "", PQV_NA, "no bank parties in the file"
        ShadeAt doc, vStart, Len(val), PQ_NA
        Exit Sub
    End If

    lineKey = Replace(NormCP(t), " ", "")
    keys = DictKeys(ef.Banks)
    For Each k In keys
        bank = CStr(k)
        share = CDbl(DictGet(ef.Banks, bank)) / ef.Total
        If share >= PQ_CP_SHARE And Not IsJunkParty(bank) Then
            AddToList bigList, bank
            If Not BankNamed(bank, lineKey) Then AddToList missing, bank
        End If
    Next k
    If missing = "" Then
        AddPQItem items, "Financial institutions", val, bigList, PQV_OK, ""
        ShadeAt doc, vStart, Len(val), PQ_OK
    Else
        AddPQItem items, "Financial institutions", val, bigList, PQV_EYE, "not named: " & missing
        ShadeAt doc, vStart, Len(val), PQ_EYE
    End If
End Sub

Private Function IsJunkParty(ByVal key As String) As Boolean
    Select Case key
        Case "", "NOTPROVIDED", "NOT PROVIDED", "NA", "N A", "UNKNOWN", "NONE"
            IsJunkParty = True
    End Select
End Function

' "COMMUNITY FEDERAL SAVINGS BANK" is named by "CFSB"; otherwise the first two words
' of the bank, spaces ignored ("JPMORGAN CHASE" matches "JP Morgan Chase")
Private Function BankNamed(ByVal bankKey As String, ByVal lineKey As String) As Boolean
    Dim words() As String, stem As String
    If Left$(bankKey, 11) = "COMMUNITY F" And InStr(lineKey, "CFSB") > 0 Then
        BankNamed = True
        Exit Function
    End If
    words = Split(bankKey, " ")
    stem = words(0)
    If UBound(words) >= 1 Then stem = stem & words(1)
    BankNamed = (InStr(lineKey, stem) > 0)
End Function

' First sentence of "Between D1 and D2, <customer> received N wire transfers totaling $X":
' rule dates, then a count/amount pair from all, credit or debit activity. Returns True
' with the written dates as the window later sentences inherit.
Private Function EscBetween(ByVal doc As Object, ByRef ctx As NarrCtx, ByRef ef As EscFacts, ByVal items As Collection, _
                            ByVal segStart As Long, ByVal t As String, ByVal sentEnd As Long, _
                            ByRef ownFrom As Date, ByRef ownTo As Date) As Boolean
    Dim dts As Collection, p As Long, n As Long, v As Long, nextP As Long, nn As Long, vv As Long
    Dim a As Long, an As Long, bound As Long
    Dim cTok As String, aTok As String, hitScope As String, cntScope As String, scopeDir As String
    Dim scopes As Variant, s As Variant
    Dim sn As Long, sAmt As Double, sAch As Long, sWire As Long, dummy As Date, hasWin As Boolean

    Set dts = SentenceDates(t, 1, sentEnd)
    If dts.Count >= 2 Then
        JudgeRuleDate doc, items, segStart + CLng(dts(1)(0)) - 1, Mid$(t, CLng(dts(1)(0)), CLng(dts(1)(1))), _
                      "Between start date", ctx.NarrFrom, ctx
        JudgeRuleDate doc, items, segStart + CLng(dts(2)(0)) - 1, Mid$(t, CLng(dts(2)(0)), CLng(dts(2)(1))), _
                      "Between end date", ctx.NarrTo, ctx
        ownFrom = dts(1)(2)
        ownTo = dts(2)(2)
        hasWin = True
        EscBetween = True
    End If

    scopes = Array(Array("all activity", ctx.NewCount, ctx.NewTotal, ""), _
                   Array("credits", ctx.NewCrCount, ctx.NewCrAmt, "C"), _
                   Array("debits", ctx.NewDrCount, ctx.NewDrAmt, "D"))
    p = NextCount(t, 1, sentEnd, Array("transaction", "transfer", "txn", "credit", "debit"), n, v)
    If p = 0 Then Exit Function
    cTok = Mid$(t, p, n)
    nextP = NextCount(t, p + n, sentEnd, Array("transaction", "transfer", "txn", "credit", "debit"), nn, vv)
    If nextP > 0 Then bound = nextP - 1 Else bound = sentEnd
    a = AmountAfterTotaling(t, p + n, bound, an)
    If a > 0 Then aTok = Mid$(t, a, an)

    For Each s In scopes
        If v = CLng(s(1)) Then
            If cntScope = "" Then
                cntScope = CStr(s(0))
                scopeDir = CStr(s(3))
            End If
            If aTok = "" Then
                hitScope = CStr(s(0))
            ElseIf SameMoney(ValueToNumber(aTok), CDbl(s(2))) Then
                hitScope = CStr(s(0))
            End If
            If hitScope <> "" Then
                scopeDir = CStr(s(3))
                Exit For
            End If
        End If
    Next s

    If hitScope <> "" Then
        MarkToken doc, items, segStart + p - 1, cTok, "Between count", UsNumber(v) & " (" & hitScope & ")", PQV_OK, ""
        If aTok <> "" Then MarkToken doc, items, segStart + a - 1, aTok, "Between amount", hitScope, PQV_OK, ""
    ElseIf cntScope <> "" Then
        MarkToken doc, items, segStart + p - 1, cTok, "Between count", UsNumber(v) & " (" & cntScope & ")", PQV_OK, ""
        If aTok <> "" Then JudgeScopeAmount doc, items, segStart + a - 1, aTok, "Between amount", scopes, cntScope
    Else
        MarkToken doc, items, segStart + p - 1, cTok, "Between count", UsNumber(ctx.NewCount), PQV_BAD, _
                  "the file has " & UsNumber(ctx.NewCount) & " (credits " & UsNumber(ctx.NewCrCount) & _
                  ", debits " & UsNumber(ctx.NewDrCount) & ")"
        If aTok <> "" Then
            If AnyScopeAmount(scopes, ValueToNumber(aTok)) Then
                MarkToken doc, items, segStart + a - 1, aTok, "Between amount", "", PQV_OK, ""
            Else
                MarkToken doc, items, segStart + a - 1, aTok, "Between amount", "$" & UsAmount(ctx.NewTotal), _
                          PQV_BAD, "the file says $" & UsAmount(ctx.NewTotal)
            End If
        End If
    End If
    If cntScope <> "" Then
        SubsetStats ef, False, dummy, dummy, "", scopeDir, sn, sAmt, sAch, sWire
        CheckInstrument doc, items, segStart, t, p + n, sn, sAch, sWire, "Between"
    End If

    ' Further claims in the same sentence
    If nextP > 0 Then
        EscSentence doc, ctx, ef, items, segStart, t, nextP, sentEnd, hasWin, ownFrom, ownTo, "Between", dummy, dummy
    End If
End Function

Private Sub JudgeScopeAmount(ByVal doc As Object, ByVal items As Collection, ByVal absPos As Long, ByVal tok As String, _
                             ByVal label As String, ByVal scopes As Variant, ByVal scopeName As String)
    Dim s As Variant, want As Double, v As Double
    For Each s In scopes
        If CStr(s(0)) = scopeName Then want = CDbl(s(2))
    Next s
    v = ValueToNumber(tok)
    If SameMoney(v, -Int(-Abs(want))) Then
        MarkToken doc, items, absPos, tok, label, "$" & UsAmount(want), PQV_EYE, "whole-dollar figure; this line uses cents"
    Else
        MarkToken doc, items, absPos, tok, label, "$" & UsAmount(want), PQV_BAD, _
                  "the file says $" & UsAmount(want) & " for " & scopeName
    End If
End Sub

Private Function AnyScopeAmount(ByVal scopes As Variant, ByVal v As Double) As Boolean
    Dim s As Variant
    For Each s In scopes
        If SameMoney(v, CDbl(s(2))) Then
            AnyScopeAmount = True
            Exit Function
        End If
    Next s
End Function

' "wire" or "ACH" right after a count, against the transaction codes of the rows described
Private Sub CheckInstrument(ByVal doc As Object, ByVal items As Collection, ByVal segStart As Long, ByVal t As String, _
                            ByVal afterPos As Long, ByVal n As Long, ByVal nAch As Long, ByVal nWire As Long, _
                            ByVal lineName As String)
    Dim claim As String, wPos As Long, wLen As Long, other As Long, same As Long, word As String, mix As String

    claim = InstrumentClaim(t, afterPos, wPos, wLen)
    If claim = "" Or wPos = 0 Then Exit Sub
    word = Mid$(t, wPos, wLen)
    mix = UsNumber(nAch) & " ACH/IAT, " & UsNumber(nWire) & " wire"
    If nAch + nWire = 0 Then
        MarkToken doc, items, segStart + wPos - 1, word, lineName & " instrument", "", PQV_NA, "no transaction codes to compare"
        Exit Sub
    End If
    If claim = "WIRE" Then
        other = nAch
        same = nWire
    Else
        other = nWire
        same = nAch
    End If
    If other = 0 Then
        MarkToken doc, items, segStart + wPos - 1, word, lineName & " instrument", mix, PQV_OK, ""
    ElseIf other >= same Then
        MarkToken doc, items, segStart + wPos - 1, word, lineName & " instrument", mix, PQV_BAD, _
                  UsNumber(other) & " of " & UsNumber(n) & IIf(claim = "WIRE", " are ACH/IAT", " are wires")
    Else
        MarkToken doc, items, segStart + wPos - 1, word, lineName & " instrument", mix, PQV_EYE, _
                  UsNumber(other) & " of " & UsNumber(n) & IIf(claim = "WIRE", " are ACH/IAT", " are wires")
    End If
End Sub

' One sentence (t, lo..hi). Every "N ... totaling $X" claim is recomputed from the rows
' in the date window and for the counterparty named in the sentence; other amounts,
' dates and counts are matched against the fact book. Returns True with the
' sentence's first date pair.
Private Function EscSentence(ByVal doc As Object, ByRef ctx As NarrCtx, ByRef ef As EscFacts, ByVal items As Collection, _
                             ByVal segStart As Long, ByVal t As String, ByVal lo As Long, ByVal hi As Long, _
                             ByVal hasInh As Boolean, ByVal inhFrom As Date, ByVal inhTo As Date, _
                             ByVal lineName As String, ByRef ownFrom As Date, ByRef ownTo As Date) As Boolean
    Dim dts As Collection, opts As Collection, handled As Object
    Dim followAll As Variant, d As Variant
    Dim p As Long, n As Long, v As Long, nextP As Long, nn As Long, vv As Long, a As Long, an As Long, bound As Long
    Dim cTok As String, aTok As String, hint As String, low As String

    If hi < lo Then Exit Function
    followAll = Array("transaction", "transfer", "txn", "credit", "debit", "payment", "wire", "deposit")
    Set dts = SentenceDates(t, lo, hi)
    If dts.Count >= 2 Then
        ownFrom = dts(1)(2)
        ownTo = dts(2)(2)
        EscSentence = True
    End If
    For Each d In dts
        If InReviewWindow(ctx, Mid$(t, CLng(d(0)), CLng(d(1)))) Then
            BindDateToken doc, ctx, ef, items, segStart + CLng(d(0)) - 1, Mid$(t, CLng(d(0)), CLng(d(1))), "Date"
        End If
    Next d

    Set opts = CPOptions(ef, Mid$(t, lo, hi - lo + 1))
    Set handled = CreateLookupDict()

    ' "N ... totaling $X" claims
    p = NextCount(t, lo, hi, followAll, n, v)
    Do While p > 0
        cTok = Mid$(t, p, n)
        nextP = NextCount(t, p + n, hi, followAll, nn, vv)
        If Not LooksLikeYear(cTok) Then
            If nextP > 0 Then bound = nextP - 1 Else bound = hi
            a = AmountAfterTotaling(t, p + n, bound, an)
            aTok = ""
            If a > 0 Then
                aTok = Mid$(t, a, an)
                If Not DictExists(handled, CStr(a)) Then DictAdd handled, CStr(a), True
            End If
            hint = CountDirection(t, p, n)
            EscClaim doc, ef, items, segStart, t, dts, opts, p, n, v, a, aTok, hint, hasInh, inhFrom, inhTo, lineName
        End If
        If nextP = 0 Then Exit Do
        p = nextP
        n = nn
        v = vv
    Loop

    ' "N counterparties"
    EscCounterpartyCounts doc, ef, items, segStart, t, lo, hi, dts, hasInh, inhFrom, inhTo

    ' Amounts not attached to a count
    EscLooseAmounts doc, ef, items, segStart, t, lo, hi, dts, opts, handled, hasInh, inhFrom, inhTo

    ' Wording the data can contradict
    low = LCase$(Replace(Mid$(t, lo, hi - lo + 1), Chr(160), " "))
    EscHighDollarClaim doc, ef, items, segStart, t, lo, low, opts
    EscConsecutiveClaim doc, ef, items, segStart, t, lo, low, opts, dts, hasInh, inhFrom, inhTo
End Function

' Candidate windows for a claim, as Array(useWindow, from, to): its own window, else
' the inherited window and then no window
Private Function ClaimWindows(ByVal hasOwn As Boolean, ByVal wFrom As Date, ByVal wTo As Date, _
                              ByVal hasInh As Boolean, ByVal inhFrom As Date, ByVal inhTo As Date) As Variant
    If hasOwn Then
        ClaimWindows = Array(Array(True, wFrom, wTo))
    ElseIf hasInh Then
        ClaimWindows = Array(Array(True, inhFrom, inhTo), Array(False, inhFrom, inhTo))
    Else
        ClaimWindows = Array(Array(False, wFrom, wTo))
    End If
End Function

Private Sub EscClaim(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, ByVal segStart As Long, _
                     ByVal t As String, ByVal dts As Collection, ByVal opts As Collection, _
                     ByVal p As Long, ByVal n As Long, ByVal v As Long, ByVal a As Long, ByVal aTok As String, _
                     ByVal hint As String, ByVal hasInh As Boolean, ByVal inhFrom As Date, ByVal inhTo As Date, _
                     ByVal lineName As String)
    Dim hasOwn As Boolean, wFrom As Date, wTo As Date, wv As Variant
    Dim cpIdx As Long, dirIdx As Long, useWin As Boolean, tf As Date, tt As Date
    Dim cpOpt As String, dirOpt As String, found As Boolean, cTok As String, label As String, note As String
    Dim sn As Long, sAmt As Double, sAch As Long, sWire As Long, aVal As Double

    cTok = Mid$(t, p, n)
    If aTok <> "" Then aVal = ValueToNumber(aTok)
    hasOwn = ClaimWindow(t, dts, p, wFrom, wTo)

    For Each wv In ClaimWindows(hasOwn, wFrom, wTo, hasInh, inhFrom, inhTo)
        useWin = CBool(wv(0))
        tf = CDate(wv(1))
        tt = CDate(wv(2))
        For cpIdx = 0 To opts.Count
            If cpIdx = 0 Then cpOpt = "" Else cpOpt = CStr(opts(cpIdx))
            For dirIdx = 0 To 1
                If dirIdx = 0 Then
                    dirOpt = ""
                Else
                    dirOpt = hint
                End If
                If dirIdx = 0 Or hint <> "" Then
                    SubsetStats ef, useWin, tf, tt, cpOpt, dirOpt, sn, sAmt, sAch, sWire
                    If sn = v And (aTok = "" Or SameMoney(sAmt, aVal)) Then
                        found = True
                        label = ScopeText(useWin, tf, tt, cpOpt, dirOpt)
                        Exit For
                    End If
                End If
            Next dirIdx
            If found Then Exit For
        Next cpIdx
        If found Then Exit For
    Next wv

    If found Then
        MarkToken doc, items, segStart + p - 1, cTok, lineName & " count", UsNumber(v) & " (" & label & ")", PQV_OK, ""
        If aTok <> "" Then MarkToken doc, items, segStart + a - 1, aTok, lineName & " amount", label, PQV_OK, ""
        CheckInstrument doc, items, segStart, t, p + n, sn, sAch, sWire, lineName
        Exit Sub
    End If

    ' No subset fits: fact book, with the most specific reading as a hint
    If hasOwn Then
        useWin = True
        tf = wFrom
        tt = wTo
    Else
        useWin = hasInh
        tf = inhFrom
        tt = inhTo
    End If
    cpOpt = ""
    If opts.Count > 0 Then cpOpt = CStr(opts(1))
    SubsetStats ef, useWin, tf, tt, cpOpt, hint, sn, sAmt, sAch, sWire
    note = "the file has " & UsNumber(sn) & " totaling $" & UsAmount(sAmt) & " for " & ScopeText(useWin, tf, tt, cpOpt, hint)
    If CntFactLabel(ef, v) <> "" Then
        MarkToken doc, items, segStart + p - 1, cTok, lineName & " count", CntFactLabel(ef, v), PQV_OK, ""
    Else
        MarkToken doc, items, segStart + p - 1, cTok, lineName & " count", "", PQV_EYE, note
    End If
    If aTok <> "" Then
        If AmtFactLabel(ef, aVal) <> "" Then
            MarkToken doc, items, segStart + a - 1, aTok, lineName & " amount", AmtFactLabel(ef, aVal), PQV_OK, ""
        Else
            MarkToken doc, items, segStart + a - 1, aTok, lineName & " amount", "", PQV_EYE, note
        End If
    End If
End Sub

Private Sub EscCounterpartyCounts(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, _
                                  ByVal segStart As Long, ByVal t As String, ByVal lo As Long, ByVal hi As Long, _
                                  ByVal dts As Collection, ByVal hasInh As Boolean, ByVal inhFrom As Date, ByVal inhTo As Date)
    Dim p As Long, n As Long, v As Long, useWin As Boolean, wFrom As Date, wTo As Date
    Dim distinct As Object, i As Long, cnt As Long

    p = NextCount(t, lo, hi, Array("counterpart"), n, v)
    Do While p > 0
        useWin = ClaimWindow(t, dts, p, wFrom, wTo)
        If Not useWin And hasInh Then
            useWin = True
            wFrom = inhFrom
            wTo = inhTo
        End If
        Set distinct = CreateLookupDict()
        For i = 1 To ef.TxN
            If ef.TxCP(i) <> "" Then
                If TxInScope(ef, i, useWin, wFrom, wTo, "", "", "", "") Then
                    If Not DictExists(distinct, ef.TxCP(i)) Then DictAdd distinct, ef.TxCP(i), True
                End If
            End If
        Next i
        cnt = DictCount(distinct)
        If v = cnt Then
            MarkToken doc, items, segStart + p - 1, Mid$(t, p, n), "Counterparty count", UsNumber(cnt), PQV_OK, ""
        Else
            MarkToken doc, items, segStart + p - 1, Mid$(t, p, n), "Counterparty count", UsNumber(cnt), PQV_EYE, _
                      "the file has " & UsNumber(cnt) & " distinct counterparties"
        End If
        p = NextCount(t, p + n, hi, Array("counterpart"), n, v)
    Loop
End Sub

Private Sub EscLooseAmounts(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, _
                            ByVal segStart As Long, ByVal t As String, ByVal lo As Long, ByVal hi As Long, _
                            ByVal dts As Collection, ByVal opts As Collection, ByVal handled As Object, _
                            ByVal hasInh As Boolean, ByVal inhFrom As Date, ByVal inhTo As Date)
    Dim pos As Long, n As Long, tok As String, v As Double, label As String
    Dim hasOwn As Boolean, wFrom As Date, wTo As Date, wv As Variant, useWin As Boolean, tf As Date, tt As Date
    Dim cpIdx As Long, dirIdx As Long, cpOpt As String, dirOpt As String
    Dim sn As Long, sAmt As Double, sAch As Long, sWire As Long

    pos = lo
    Do
        pos = FindAmountToken(t, pos, hi, n)
        If pos = 0 Then Exit Do
        If Not DictExists(handled, CStr(pos)) Then
            tok = Mid$(t, pos, n)
            v = ValueToNumber(tok)
            label = ""
            hasOwn = ClaimWindow(t, dts, pos, wFrom, wTo)
            For Each wv In ClaimWindows(hasOwn, wFrom, wTo, hasInh, inhFrom, inhTo)
                useWin = CBool(wv(0))
                tf = CDate(wv(1))
                tt = CDate(wv(2))
                For cpIdx = 0 To opts.Count
                    If cpIdx = 0 Then cpOpt = "" Else cpOpt = CStr(opts(cpIdx))
                    For dirIdx = 0 To 2
                        If dirIdx = 0 Then
                            dirOpt = ""
                        ElseIf dirIdx = 1 Then
                            dirOpt = "C"
                        Else
                            dirOpt = "D"
                        End If
                        If useWin Or cpOpt <> "" Or dirOpt <> "" Then
                            SubsetStats ef, useWin, tf, tt, cpOpt, dirOpt, sn, sAmt, sAch, sWire
                            If sn > 0 And SameMoney(sAmt, v) Then
                                label = ScopeText(useWin, tf, tt, cpOpt, dirOpt)
                                Exit For
                            End If
                        End If
                    Next dirIdx
                    If label <> "" Then Exit For
                Next cpIdx
                If label <> "" Then Exit For
            Next wv
            If label = "" Then label = AmtFactLabel(ef, v)
            BindFactToken doc, items, segStart + pos - 1, tok, "Amount", label, "no figure in the file matches this"
        End If
        pos = pos + n
    Loop
End Sub

' "All transactions were for high dollar amounts" against the smallest transactions
Private Sub EscHighDollarClaim(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, _
                               ByVal segStart As Long, ByVal t As String, ByVal lo As Long, ByVal low As String, _
                               ByVal opts As Collection)
    Dim i As Long, nSmall As Long, smallest As Double, cpOpt As String, k As String, fam As String
    Dim p As Long, word As String, dummy As Date

    If Not ((" " & low) Like "* all *transaction*" Or (" " & low) Like "* all *transfer*") Then Exit Sub
    If Not (low Like "*high*dollar*" Or low Like "*high*value*") Then Exit Sub
    If opts.Count > 0 Then
        cpOpt = CStr(opts(1))
        k = Mid$(cpOpt, 2)
        fam = k & " "
    End If
    For i = 1 To ef.TxN
        If TxInScope(ef, i, False, dummy, dummy, cpOpt, k, fam, "") Then
            If ef.TxAmt(i) < PQ_SMALL_AMT Then
                nSmall = nSmall + 1
                If smallest = 0 Or ef.TxAmt(i) < smallest Then smallest = ef.TxAmt(i)
            End If
        End If
    Next i
    If nSmall = 0 Then Exit Sub
    p = InStr(1, Mid$(t, lo), "high", vbTextCompare)
    If p > 0 Then
        p = lo + p - 1
        word = Mid$(t, p, 4)
        MarkToken doc, items, segStart + p - 1, word, "High-dollar wording", "", PQV_EYE, _
                  UsNumber(nSmall) & " transactions are under $" & UsWholeAmount(PQ_SMALL_AMT) & _
                  " (smallest $" & UsAmount(smallest) & ")"
    End If
End Sub

' "... on nearly consecutive days" against the gaps between transaction dates
Private Sub EscConsecutiveClaim(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, _
                                ByVal segStart As Long, ByVal t As String, ByVal lo As Long, ByVal low As String, _
                                ByVal opts As Collection, ByVal dts As Collection, _
                                ByVal hasInh As Boolean, ByVal inhFrom As Date, ByVal inhTo As Date)
    Dim useWin As Boolean, wFrom As Date, wTo As Date, cpOpt As String, k As String, fam As String
    Dim days As Object, keys As Variant, i As Long, j As Long, tmp As Long, gap As Long
    Dim dayNums() As Long, nDays As Long, closeGaps As Long, total As Long
    Dim gapCount As Object, bestGap As Long, bestN As Long, p As Long

    If InStr(low, "consecutive") = 0 Then Exit Sub
    If InStr(ef.OnceFlags, "|consecutive|") > 0 Then Exit Sub
    If dts.Count >= 2 Then
        useWin = True
        wFrom = dts(1)(2)
        wTo = dts(2)(2)
    ElseIf hasInh Then
        useWin = True
        wFrom = inhFrom
        wTo = inhTo
    End If
    If opts.Count > 0 Then
        cpOpt = CStr(opts(1))
        k = Mid$(cpOpt, 2)
        fam = k & " "
    End If

    Set days = CreateLookupDict()
    For i = 1 To ef.TxN
        If ef.TxHasDate(i) Then
            If TxInScope(ef, i, useWin, wFrom, wTo, cpOpt, k, fam, "") Then
                If Not DictExists(days, CStr(CLng(ef.TxDate(i)))) Then DictAdd days, CStr(CLng(ef.TxDate(i))), True
            End If
        End If
    Next i
    nDays = DictCount(days)
    If nDays < 4 Then Exit Sub
    keys = DictKeys(days)
    ReDim dayNums(0 To nDays - 1)
    For i = 0 To nDays - 1
        dayNums(i) = CLng(keys(i))
    Next i
    For i = 0 To nDays - 2
        For j = i + 1 To nDays - 1
            If dayNums(j) < dayNums(i) Then
                tmp = dayNums(i)
                dayNums(i) = dayNums(j)
                dayNums(j) = tmp
            End If
        Next j
    Next i
    Set gapCount = CreateLookupDict()
    For i = 1 To nDays - 1
        gap = dayNums(i) - dayNums(i - 1)
        total = total + 1
        If gap <= 2 Then closeGaps = closeGaps + 1
        AddToAmount gapCount, CStr(gap), 1
        If CLng(DictGet(gapCount, CStr(gap))) > bestN Then
            bestN = CLng(DictGet(gapCount, CStr(gap)))
            bestGap = gap
        End If
    Next i
    If closeGaps / total >= PQ_CLOSE_GAP_SHARE Then Exit Sub

    ef.OnceFlags = ef.OnceFlags & "|consecutive|"
    p = InStr(1, Mid$(t, lo), "consecutive", vbTextCompare)
    If p > 0 Then
        p = lo + p - 1
        MarkToken doc, items, segStart + p - 1, Mid$(t, p, 11), "Consecutive-days wording", _
                  ScopeText(useWin, wFrom, wTo, cpOpt, ""), PQV_EYE, _
                  UsNumber(closeGaps) & " of " & UsNumber(total) & " gaps between transaction dates are 1-2 days; " & _
                  "the most common gap is " & bestGap & " days (" & bestN & " times)"
    End If
End Sub

' Dates of birth, incorporation dates and the like fall outside the review window
' and are not transaction claims
Private Function InReviewWindow(ByRef ctx As NarrCtx, ByVal tok As String) As Boolean
    Dim d As Date, ok As Boolean, lo As Variant, hi As Variant
    d = ParseDateValue(tok, ok)
    If Not ok Then Exit Function
    lo = ctx.NarrFrom
    hi = ctx.NarrTo
    If Not IsDate(lo) Then lo = ctx.NewMin
    If Not IsDate(hi) Then hi = ctx.NewMax
    If Not IsDate(lo) Or Not IsDate(hi) Then Exit Function
    InReviewWindow = (d >= CDate(lo) - 31 And d <= CDate(hi) + 31)
End Function

Private Sub BindDateToken(ByVal doc As Object, ByRef ctx As NarrCtx, ByRef ef As EscFacts, ByVal items As Collection, _
                          ByVal absPos As Long, ByVal tok As String, ByVal label As String)
    Dim d As Date, ok As Boolean, lbl As String
    d = ParseDateValue(tok, ok)
    If ok Then
        If DictExists(ef.DateSet, Format$(d, "yyyymmdd")) Then
            lbl = "transaction date"
        ElseIf SameDay(d, ctx.NarrFrom) Or SameDay(d, ctx.NarrTo) Then
            lbl = "review period"
        End If
    End If
    BindFactToken doc, items, absPos, tok, label, lbl, "no transaction on this date"
End Sub

Private Function SameDay(ByVal d As Date, ByVal v As Variant) As Boolean
    If Not IsDate(v) Then Exit Function
    SameDay = (Int(CDbl(d)) = Int(CDbl(CDate(v))))
End Function

Private Sub BindFactToken(ByVal doc As Object, ByVal items As Collection, ByVal absPos As Long, ByVal tok As String, _
                          ByVal label As String, ByVal factLabel As String, ByVal missNote As String)
    If factLabel <> "" Then
        MarkToken doc, items, absPos, tok, label, factLabel, PQV_OK, ""
    Else
        MarkToken doc, items, absPos, tok, label, "", PQV_EYE, missNote
    End If
End Sub

Private Sub MarkToken(ByVal doc As Object, ByVal items As Collection, ByVal absPos As Long, ByVal tok As String, _
                      ByVal label As String, ByVal fileVal As String, ByVal verdict As String, ByVal note As String)
    AddPQItem items, label, tok, fileVal, verdict, note
    Select Case verdict
        Case PQV_OK
            ShadeAt doc, absPos, Len(tok), PQ_OK
        Case PQV_BAD
            ShadeAt doc, absPos, Len(tok), PQ_BAD
        Case PQV_EYE
            ShadeAt doc, absPos, Len(tok), PQ_EYE
        Case Else
            ShadeAt doc, absPos, Len(tok), PQ_NA
    End Select
End Sub

' Whole-dollar figures (opening sentence, Total Suspicious Dollar Amount): the total rounded up
Private Sub JudgeWhole(ByVal doc As Object, ByVal items As Collection, ByVal absPos As Long, ByVal tok As String, _
                       ByVal label As String, ByVal total As Double)
    Dim v As Double, want As String
    v = ValueToNumber(tok)
    want = "$" & UsWholeAmount(total)
    If SameMoney(v, -Int(-Abs(total))) Then
        MarkToken doc, items, absPos, tok, label, want, PQV_OK, ""
    ElseIf Abs(v - Abs(total)) < 1 Then
        MarkToken doc, items, absPos, tok, label, want, PQV_EYE, "within $1 of the total; the rule is whole dollars rounded up"
    Else
        MarkToken doc, items, absPos, tok, label, want, PQV_BAD, "the file says " & want
    End If
End Sub

' Dates governed by NARRATIVE_DATE_RANGE: the file period, or the first/last
' transaction when the file name has no period
Private Sub JudgeRuleDate(ByVal doc As Object, ByVal items As Collection, ByVal absPos As Long, ByVal tok As String, _
                          ByVal label As String, ByVal want As Variant, ByRef ctx As NarrCtx)
    Dim d As Date, ok As Boolean
    If Not IsDate(want) Then
        MarkToken doc, items, absPos, tok, label, "", PQV_NA, "the file has no date to compare"
        Exit Sub
    End If
    d = ParseDateValue(tok, ok)
    If ok And SameDay(d, want) Then
        MarkToken doc, items, absPos, tok, label, UsDate(want), PQV_OK, ""
    ElseIf ok And (SameDay(d, ctx.NewMin) Or SameDay(d, ctx.NewMax)) Then
        MarkToken doc, items, absPos, tok, label, UsDate(want), PQV_EYE, _
                  "matches the first/last transaction; the rule is the file period " & UsDate(ctx.NarrFrom) & _
                  " to " & UsDate(ctx.NarrTo)
    Else
        MarkToken doc, items, absPos, tok, label, UsDate(want), PQV_BAD, "the file says " & UsDate(want)
    End If
End Sub

' Transfers from the customer's own name that the narrative does not describe
Private Sub EscCounterparties(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection)
    Dim docText As String, keys As Variant, k As Variant
    Dim ownAmt As Double, ownShare As Double, ownN As Long, i As Long
    Dim own As Object

    If DictCount(ef.CPs) = 0 Or ef.Total = 0 Then Exit Sub
    docText = doc.Content.Text
    Set own = CreateLookupDict()
    keys = DictKeys(ef.CPs)

    If ef.Holder <> "" Then
        For Each k In keys
            If SameParty(CStr(k), ef.Holder) Then
                DictAdd own, CStr(k), True
                If DictExists(ef.CPAmt, CStr(k)) Then ownAmt = ownAmt + CDbl(DictGet(ef.CPAmt, CStr(k)))
            End If
        Next k
        ownShare = ownAmt / ef.Total
        If ownShare >= PQ_CP_SHARE And ownShare < PQ_OWN_MAX_SHARE And Not MentionsOwnTransfers(docText) Then
            For i = 1 To ef.TxN
                If DictExists(own, ef.TxCP(i)) Then ownN = ownN + 1
            Next i
            AddPQItem items, "Own-name transfers", "", UsNumber(ownN) & " totaling $" & UsAmount(ownAmt) & _
                      " (" & Format$(ownShare * 100, "0") & "%)", PQV_EYE, _
                      "transfers from the customer's own name are not described"
        End If
    End If

End Sub

Private Function MentionsOwnTransfers(ByVal docText As String) As Boolean
    Dim s As String
    s = LCase$(docText)
    MentionsOwnTransfers = (InStr(s, "self") > 0 Or InStr(s, "own account") > 0 Or _
                            InStr(s, "between its accounts") > 0 Or InStr(s, "between their accounts") > 0 Or _
                            InStr(s, "between his accounts") > 0 Or InStr(s, "between her accounts") > 0)
End Function

' Alerted transaction file: every alerted row must be in the lookback file, the lookback
' flag should agree, and the narrative should describe the alerted activity
Private Sub EscAlerted(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection)
    Dim docText As String, mentioned As Boolean

    If Not ef.HasAlertFile Then Exit Sub
    If ef.AlertMissing > 0 Then
        AddPQItem items, "Alerted transactions", ef.AlertFile, UsNumber(ef.AlertMissing) & " missing", PQV_BAD, _
                  "alerted transactions that are not in the lookback file"
    End If
    If ef.LookbackAlerted <> ef.AlertN Then
        AddPQItem items, "Lookback alerted flag", "", "lookback flags " & UsNumber(ef.LookbackAlerted) & _
                  ", the alerted file has " & UsNumber(ef.AlertN), PQV_EYE, "data quality"
    End If
    docText = doc.Content.Text
    If ef.AlertN > 0 Then
        mentioned = (InStr(docText, UsAmount(ef.AlertTotal)) > 0 Or InStr(docText, UsWholeAmount(ef.AlertTotal)) > 0)
        If Not mentioned And ef.AlertRule <> "" Then mentioned = (InStr(1, docText, ef.AlertRule, vbTextCompare) > 0)
        If Not mentioned Then
            AddPQItem items, "Alerted activity", "", UsNumber(ef.AlertN) & " totaling $" & UsAmount(ef.AlertTotal) & ", " & _
                      DateText(ef.AlertMin) & " - " & DateText(ef.AlertMax) & IIf(ef.AlertRule <> "", ", " & ef.AlertRule, ""), _
                      PQV_EYE, "the alerted activity is not described"
        End If
    End If
End Sub

' Words one or two letters away from the customer's name ("Midsagency" for Midagency)
Private Sub EscNameCheck(ByVal doc As Object, ByRef ef As EscFacts, ByVal items As Collection, ByVal customer As String)
    Dim target As String, src As Variant, d As Collection, w As Variant, maxD As Long
    Dim docText As String, i As Long, j As Long, word As String, uw As String, dist As Long
    Dim variants As Object, known As Object, keys As Variant, k As Variant, hits As Collection, h As Variant, shaded As Long

    For Each src In Array(ef.Holder, NormCP(customer))
        Set d = Distinctive(CStr(src))
        For Each w In d
            If Len(CStr(w)) >= 5 And Len(CStr(w)) > Len(target) Then target = CStr(w)
        Next w
        If target <> "" Then Exit For
    Next src
    If target = "" Then Exit Sub
    If Len(target) <= 6 Then maxD = 1 Else maxD = 2

    ' Words that belong to counterparties or banks are not misspellings of the customer
    Set known = CreateLookupDict()
    For Each src In Array(ef.CPs, ef.Banks)
        If DictCount(src) > 0 Then
            keys = DictKeys(src)
            For Each k In keys
                For Each w In Split(CStr(k), " ")
                    If CStr(w) <> "" Then
                        If Not DictExists(known, CStr(w)) Then DictAdd known, CStr(w), True
                    End If
                Next w
            Next k
        End If
    Next src

    Set variants = CreateLookupDict()
    docText = doc.Content.Text
    i = 1
    Do While i <= Len(docText)
        If Mid$(docText, i, 1) Like "[A-Za-z]" Then
            j = i
            Do While j <= Len(docText)
                If Not (Mid$(docText, j, 1) Like "[A-Za-z]") Then Exit Do
                j = j + 1
            Loop
            word = Mid$(docText, i, j - i)
            If Len(word) >= 5 Then
                uw = UCase$(word)
                If uw <> target And Left$(uw, 1) = Left$(target, 1) And Not DictExists(known, uw) Then
                    dist = Levenshtein(uw, target)
                    If dist >= 1 And dist <= maxD Then AddToAmount variants, word, 1
                End If
            End If
            i = j
        Else
            i = i + 1
        End If
    Loop

    If DictCount(variants) = 0 Then Exit Sub
    keys = DictKeys(variants)
    For Each k In keys
        AddPQItem items, "Name spelling", CStr(k), target, PQV_EYE, "used " & CLng(DictGet(variants, CStr(k))) & _
                  " time(s); the account holder is " & IIf(ef.Holder <> "", ef.Holder, customer)
        Set hits = WdFindIn(doc, doc.Content.Start, doc.Content.End, CStr(k), True)
        shaded = 0
        For Each h In hits
            ShadeExact doc, CLng(h(0)), CLng(h(1)) - CLng(h(0)), PQ_EYE
            shaded = shaded + 1
            If shaded >= 60 Then Exit For
        Next h
    Next k
End Sub

' =========================================================================
' [PQC] ESCALATION: COUNTERPARTY SELECTION
' =========================================================================
' The counterparties an escalation narrative is expected to cover:
'   1. for each rule in the alerted transactions (largest first), the counterparty
'      with the highest dollar value on that rule - each counterparty once
'   2. then the counterparties with the highest dollar value in the lookback
'      activity that runs in the same direction as the alerted transactions
'      (credits for "many to one" and incoming rules, debits for outgoing and card rules)
' up to PQ_ESC_CP_COUNT in total. The customer's own name, and addresses or IDs sitting
' in the counterparty column, are not counterparties. Card descriptors of one merchant
' ("THE HOME DEPOT 4402" / "THE HOME DEPOT 4413") and name variants sharing a first
' word ("AMAZON.C1BWFDWAB" / "AMAZON MEXICO SERVICES INC") count as one counterparty.
Private Sub EscSelection(ByVal doc As Object, ByRef ef As EscFacts, ByRef aef As EscFacts, ByVal items As Collection)
    Dim stems As Object, eligible As Object, ruleAmt As Object, ruleUnitAmt As Object
    Dim dirs As Object, poolAmt As Object, chosenSet As Object, ruleUsed As Object
    Dim chosen As Collection
    Dim useAlertFile As Boolean, i As Long, n As Long, u As String, bestRule As String, bestUnit As String
    Dim amt As Double, bestAmt As Double, keys As Variant, k As Variant, prefix As String
    Dim docKey As String, docWords As Object, entry As Variant, named As Boolean, display As String

    Set stems = CPStems(ef, aef)
    Set eligible = CreateLookupDict()
    Set ruleAmt = CreateLookupDict()
    Set ruleUnitAmt = CreateLookupDict()
    Set dirs = CreateLookupDict()
    Set poolAmt = CreateLookupDict()
    Set chosenSet = CreateLookupDict()
    Set ruleUsed = CreateLookupDict()
    Set chosen = New Collection

    ' 1. alerted transactions, by rule
    useAlertFile = (aef.TxN > 0)
    If useAlertFile Then n = aef.TxN Else n = ef.TxN
    For i = 1 To n
        If useAlertFile Then
            AddAlertedRow ef, stems, eligible, ruleAmt, ruleUnitAmt, dirs, aef.TxCP(i), aef.TxAmt(i), aef.TxDir(i), aef.TxRule(i)
        ElseIf ef.TxAl(i) Then
            AddAlertedRow ef, stems, eligible, ruleAmt, ruleUnitAmt, dirs, ef.TxCP(i), ef.TxAmt(i), ef.TxDir(i), ef.TxRule(i)
        End If
    Next i

    Do
        bestRule = ""
        bestAmt = -1
        If DictCount(ruleAmt) > 0 Then
            keys = DictKeys(ruleAmt)
            For Each k In keys
                If Not DictExists(ruleUsed, CStr(k)) Then
                    If CDbl(DictGet(ruleAmt, CStr(k))) > bestAmt Then
                        bestAmt = CDbl(DictGet(ruleAmt, CStr(k)))
                        bestRule = CStr(k)
                    End If
                End If
            Next k
        End If
        If bestRule = "" Then Exit Do
        DictAdd ruleUsed, bestRule, True

        bestUnit = ""
        bestAmt = -1
        prefix = bestRule & "|"
        If DictCount(ruleUnitAmt) > 0 Then
            keys = DictKeys(ruleUnitAmt)
            For Each k In keys
                If Left$(CStr(k), Len(prefix)) = prefix Then
                    u = Mid$(CStr(k), Len(prefix) + 1)
                    If Not DictExists(chosenSet, u) Then
                        If CDbl(DictGet(ruleUnitAmt, CStr(k))) > bestAmt Then
                            bestAmt = CDbl(DictGet(ruleUnitAmt, CStr(k)))
                            bestUnit = u
                        End If
                    End If
                End If
            Next k
        End If
        If bestUnit <> "" Then
            DictAdd chosenSet, bestUnit, True
            chosen.Add Array(bestUnit, "top counterparty for rule '" & bestRule & "' in the alerted transactions ($" & _
                             UsAmount(bestAmt) & ")")
        End If
        If chosen.Count >= PQ_ESC_CP_COUNT Then Exit Do
    Loop

    ' 2. lookback activity in the same direction as the alerted transactions
    For i = 1 To ef.TxN
        If DictCount(dirs) = 0 Or DictExists(dirs, ef.TxDir(i)) Then
            u = SelectionUnit(ef, stems, eligible, ef.TxCP(i))
            If u <> "" Then AddToAmount poolAmt, u, ef.TxAmt(i)
        End If
    Next i
    Do While chosen.Count < PQ_ESC_CP_COUNT
        bestUnit = ""
        bestAmt = -1
        If DictCount(poolAmt) > 0 Then
            keys = DictKeys(poolAmt)
            For Each k In keys
                If Not DictExists(chosenSet, CStr(k)) Then
                    amt = CDbl(DictGet(poolAmt, CStr(k)))
                    If amt > bestAmt Then
                        bestAmt = amt
                        bestUnit = CStr(k)
                    End If
                End If
            Next k
        End If
        If bestUnit = "" Then Exit Do
        DictAdd chosenSet, bestUnit, True
        chosen.Add Array(bestUnit, "next largest in the lookback " & DirectionText(dirs) & " ($" & UsAmount(bestAmt) & ")")
    Loop

    ' Is each one described in the narrative?
    If chosen.Count = 0 Then Exit Sub
    docKey = NormCP(doc.Content.Text)
    Set docWords = WordSet(docKey)
    i = 0
    For Each entry In chosen
        i = i + 1
        u = CStr(entry(0))
        named = UnitNamed(ef, aef, stems, u, docWords, docKey)
        display = UnitDisplay(ef, aef, stems, u)
        If named Then
            AddPQItem items, "Selected counterparty " & i, display, CStr(entry(1)), PQV_OK, ""
            ShadeFirst doc, UnitFindText(u), PQ_OK
        Else
            AddPQItem items, "Selected counterparty " & i, display, CStr(entry(1)), PQV_EYE, _
                      "selected by the counterparty method but not described in the narrative"
        End If
    Next entry
End Sub

Private Sub AddAlertedRow(ByRef ef As EscFacts, ByVal stems As Object, ByVal eligible As Object, ByVal ruleAmt As Object, _
                          ByVal ruleUnitAmt As Object, ByVal dirs As Object, ByVal cpKey As String, ByVal amt As Double, _
                          ByVal dirKey As String, ByVal ruleText As String)
    Dim rules As Collection, rl As Variant, u As String
    Set rules = SplitRules(ruleText)
    If rules.Count = 0 Then rules.Add "(no rule)"
    If dirKey <> "" Then
        If Not DictExists(dirs, dirKey) Then DictAdd dirs, dirKey, True
    End If
    u = SelectionUnit(ef, stems, eligible, cpKey)
    For Each rl In rules
        AddToAmount ruleAmt, CStr(rl), amt
        If u <> "" Then AddToAmount ruleUnitAmt, CStr(rl) & "|" & u, amt
    Next rl
End Sub

' "Rule A, Rule B" on one row: a comma followed by a capitalised word starts a new rule
Private Function SplitRules(ByVal s As String) As Collection
    Dim col As New Collection, i As Long, start As Long, part As String
    Set SplitRules = col
    s = Trim$(s)
    If s = "" Then Exit Function
    start = 1
    For i = 1 To Len(s) - 2
        If Mid$(s, i, 2) = ", " And Mid$(s, i + 2, 1) Like "[A-Z]" Then
            part = Trim$(Mid$(s, start, i - start))
            If part <> "" Then col.Add part
            start = i + 2
        End If
    Next i
    part = Trim$(Mid$(s, start))
    If part <> "" Then col.Add part
End Function

' The counterparty a row counts towards, or "" for the customer's own name, addresses and IDs
Private Function SelectionUnit(ByRef ef As EscFacts, ByVal stems As Object, ByVal eligible As Object, _
                               ByVal cpKey As String) As String
    Dim st As String, ok As Boolean
    If cpKey = "" Then Exit Function
    If DictExists(eligible, cpKey) Then
        ok = CBool(DictGet(eligible, cpKey))
    Else
        ok = Not IsJunkCP(cpKey)
        If ok And ef.Holder <> "" Then ok = Not SameParty(cpKey, ef.Holder)
        DictAdd eligible, cpKey, ok
    End If
    If Not ok Then Exit Function
    st = CPStem(cpKey)
    If st <> "" Then
        If DictExists(stems, st) Then
            If CDbl(DictGet(stems, st)) >= 2 Then
                SelectionUnit = "*" & st
                Exit Function
            End If
        End If
    End If
    SelectionUnit = CPAlias(cpKey)
End Function

' Stem -> number of distinct counterparty names that share it, over both files
Private Function CPStems(ByRef ef As EscFacts, ByRef aef As EscFacts) As Object
    Dim seen As Object, stems As Object, keys As Variant, k As Variant, st As String, src As Variant
    Set seen = CreateLookupDict()
    Set stems = CreateLookupDict()
    For Each src In Array(ef.CPs, aef.CPs)
        If Not src Is Nothing Then
            If DictCount(src) > 0 Then
                keys = DictKeys(src)
                For Each k In keys
                    If Not DictExists(seen, CStr(k)) Then
                        DictAdd seen, CStr(k), True
                        st = CPStem(CStr(k))
                        If st <> "" Then AddToAmount stems, st, 1
                    End If
                Next k
            End If
        End If
    Next src
    Set CPStems = stems
End Function

' Common card descriptors written the way narratives write them
Private Function CPAlias(ByVal key As String) As String
    Dim k As String
    k = " " & key & " "
    k = Replace(k, " FACEBK ", " FACEBOOK ")
    k = Replace(k, " AMZN ", " AMAZON ")
    k = Replace(k, " MSFT ", " MICROSOFT ")
    k = Replace(k, " WAL MART ", " WALMART ")
    k = Replace(k, " WM SUPERCENTER ", " WALMART SUPERCENTER ")
    CPAlias = Trim$(k)
End Function

' First real word of a counterparty: store numbers, IDs and processor prefixes skipped
' ("THE HOME DEPOT 4402" -> HOME, "AMAZON.C1BWFDWAB" -> AMAZON, "SQ *DOUGLAS LLC" -> DOUGLAS)
Private Function CPStem(ByVal key As String) As String
    Dim parts() As String, i As Long, w As String
    parts = Split(CPAlias(key), " ")
    For i = 0 To UBound(parts)
        w = parts(i)
        If w <> "" And Not (w Like "*#*") Then
            Select Case w
                Case "THE", "SQ", "TST", "PY", "SP", "PP", "PAYPAL", "POS", "ACH", "DD", "WWW", "CASH", "APP"
                Case Else
                    If Len(w) >= 4 Then CPStem = w
                    Exit Function
            End Select
        End If
    Next i
End Function

' Addresses ("102 S E DORIAN AVE") and IDs ("CB001ACE 846C 4D69 ...") in the counterparty column
Private Function IsJunkCP(ByVal key As String) As Boolean
    Dim parts() As String, i As Long
    If Distinctive(key).Count = 0 Then
        IsJunkCP = True
        Exit Function
    End If
    parts = Split(key, " ")
    If Not IsNumeric(parts(0)) Then Exit Function
    For i = 1 To UBound(parts)
        Select Case parts(i)
            Case "AVE", "AVENUE", "ST", "STREET", "RD", "ROAD", "BLVD", "DR", "DRIVE", "LN", "LANE", "WAY", "CT", "HWY", "PKWY"
                IsJunkCP = True
                Exit Function
        End Select
    Next i
End Function

Private Function UnitNamed(ByRef ef As EscFacts, ByRef aef As EscFacts, ByVal stems As Object, ByVal u As String, _
                           ByVal docWords As Object, ByVal docKey As String) As Boolean
    Dim st As String, src As Variant, keys As Variant, k As Variant
    If Left$(u, 1) <> "*" Then
        UnitNamed = NamedIn(u, docWords, docKey)
        Exit Function
    End If
    st = Mid$(u, 2)
    If DictExists(docWords, st) Then
        UnitNamed = True
        Exit Function
    End If
    For Each src In Array(ef.CPs, aef.CPs)
        If Not src Is Nothing Then
            If DictCount(src) > 0 Then
                keys = DictKeys(src)
                For Each k In keys
                    If CPStem(CStr(k)) = st Then
                        If NamedIn(CPAlias(CStr(k)), docWords, docKey) Then
                            UnitNamed = True
                            Exit Function
                        End If
                    End If
                Next k
            End If
        End If
    Next src
End Function

Private Function UnitDisplay(ByRef ef As EscFacts, ByRef aef As EscFacts, ByVal stems As Object, ByVal u As String) As String
    Dim st As String
    If Left$(u, 1) = "*" Then
        st = Mid$(u, 2)
        UnitDisplay = st & " (" & CLng(DictGet(stems, st)) & " name variants)"
    ElseIf DictExists(ef.CPs, u) Then
        UnitDisplay = CStr(DictGet(ef.CPs, u))
    ElseIf DictExists(aef.CPs, u) Then
        UnitDisplay = CStr(DictGet(aef.CPs, u))
    Else
        UnitDisplay = u
    End If
End Function

' Text to shade for a selected counterparty: the family word, or the name's first real word
Private Function UnitFindText(ByVal u As String) As String
    Dim d As Collection
    If Left$(u, 1) = "*" Then
        UnitFindText = Mid$(u, 2)
    Else
        Set d = Distinctive(u)
        If d.Count > 0 Then UnitFindText = CStr(d(1)) Else UnitFindText = u
    End If
End Function

Private Function DirectionText(ByVal dirs As Object) As String
    Dim s As String
    s = "activity"
    If DictCount(dirs) = 1 Then
        If DictExists(dirs, "C") Then
            s = "credits"
        ElseIf DictExists(dirs, "D") Then
            s = "debits"
        End If
    End If
    DirectionText = s
End Function

Private Sub WriteEscDash(ByVal ws As Worksheet, ByVal rowIdx As Long, ByRef ctx As NarrCtx, _
                         ByVal ecmID As String, ByVal alertID As String, ByVal alertInfo As String, _
                         ByVal nOK As Long, ByVal nBad As Long, ByVal nEye As Long, _
                         ByVal docPath As String, ByVal outPath As String)
    Dim status As String

    If docPath = "" Then
        status = "No narrative found"
    ElseIf nBad > 0 Then
        status = "Mismatch - review"
    ElseIf nEye > 0 Then
        status = "Check flagged items"
    Else
        status = "Clean"
    End If

    ws.Cells(rowIdx, 2).Value = ecmID
    ws.Cells(rowIdx, 3).Value = alertID
    ws.Cells(rowIdx, 4).Value = ctx.NewFile
    ws.Cells(rowIdx, 5).Value = ctx.NewCount
    ws.Cells(rowIdx, 6).Value = ctx.NewTotal
    ws.Cells(rowIdx, 7).Value = DateText(ctx.NarrFrom) & " - " & DateText(ctx.NarrTo)
    ws.Cells(rowIdx, 8).Value = alertInfo
    ws.Cells(rowIdx, 9).Value = nOK
    ws.Cells(rowIdx, 10).Value = nBad
    ws.Cells(rowIdx, 11).Value = nEye
    ws.Cells(rowIdx, 12).Value = status
    ws.Cells(rowIdx, 13).Value = GetFileName(docPath)
    ws.Cells(rowIdx, 14).Value = GetFileName(outPath)
    ws.Cells(rowIdx, 6).NumberFormat = "$#,##0.00"

    If nBad > 0 Then
        HighlightCell ws.Cells(rowIdx, 12), COLOR_ALERT_RED, COLOR_FILL_RED
        HighlightCell ws.Cells(rowIdx, 10), COLOR_ALERT_RED, COLOR_FILL_RED
    ElseIf nEye > 0 Or docPath = "" Then
        HighlightCell ws.Cells(rowIdx, 12), COLOR_AMBER, COLOR_FILL_AMBER
    Else
        HighlightCell ws.Cells(rowIdx, 12), COLOR_SUCCESS_GREEN, COLOR_FILL_GREEN
    End If
End Sub

' =========================================================================
' [PQC] SPELLING AND GRAMMAR (Word's own proofing tools)
' =========================================================================
' Runs on the reviewed copy after the figure checks. Word lists its spelling and
' grammar errors to a macro and gives spelling suggestions, but it gives a macro no
' correction for a grammar error. So:
'   - a lowercase misspelling whose best suggestion is one clear typing slip away
'     (recieved, identifed, transfered) is corrected and shaded yellow
'   - a repeated word, a missing space after a full stop, extra spaces between words
'     and a space before a comma are corrected
'   - any other misspelling gets a wavy red underline, and each of Word's grammar
'     flags a wavy blue underline; the reviewer rewords those
' Every correction and flag is listed on the findings sheet with the verdict
' "Proofing". Capitalised and all-capital words (names, companies, codes) and words
' from the transaction file are never changed or flagged. The original narrative is
' not touched.
' =========================================================================
Private Sub ProofReviewedCopy(ByVal doc As Object, ByVal knownText As String, ByVal items As Collection)
    Dim opt As Object, saved As Variant, errDesc As String, nFlag As Long, known As Object
    Dim redMarks As Collection, m As Variant

    If Not PQ_PROOF Then Exit Sub
    On Error GoTo Failed
    Set opt = doc.Application.Options
    saved = Array(opt.CheckSpellingAsYouType, opt.CheckGrammarAsYouType, opt.CheckGrammarWithSpelling, _
                  opt.IgnoreUppercase, opt.IgnoreMixedDigits, opt.IgnoreInternetAndFileAddresses)
    opt.CheckSpellingAsYouType = True
    opt.CheckGrammarAsYouType = True
    opt.CheckGrammarWithSpelling = True
    opt.IgnoreUppercase = True
    opt.IgnoreMixedDigits = True
    opt.IgnoreInternetAndFileAddresses = True
    doc.Content.NoProofing = False
    doc.Content.LanguageID = 1033                    ' wdEnglishUS
    doc.SpellingChecked = False
    doc.GrammarChecked = False

    Set known = KnownWords(knownText & " " & PQ_PROOF_WORDS)
    If PQ_PROOF_FIX Then FixMechanical doc, items
    Set redMarks = New Collection
    ProofSpelling doc, known, items, nFlag, redMarks
    ProofGrammar doc, items, nFlag
    ' misspellings last, so a grammar flag over the whole sentence does not hide them
    For Each m In redMarks
        MarkProof doc.Range(CLng(m(0)), CLng(m(1))), WD_COLOR_RED
    Next m
    doc.ShowSpellingErrors = True
    doc.ShowGrammaticalErrors = True
    RestoreProofOptions opt, saved
    Exit Sub

Failed:
    errDesc = Err.Description
    On Error Resume Next
    If IsArray(saved) Then RestoreProofOptions opt, saved
    On Error GoTo 0
    AddPQItem items, "Spelling / grammar", "", "", PQV_PROOF, "Word's proofing tools could not be used (" & errDesc & ")"
End Sub

' Word's options are the user's own settings: put them back as they were
Private Sub RestoreProofOptions(ByVal opt As Object, ByVal saved As Variant)
    On Error Resume Next
    opt.CheckSpellingAsYouType = saved(0)
    opt.CheckGrammarAsYouType = saved(1)
    opt.CheckGrammarWithSpelling = saved(2)
    opt.IgnoreUppercase = saved(3)
    opt.IgnoreMixedDigits = saved(4)
    opt.IgnoreInternetAndFileAddresses = saved(5)
    On Error GoTo 0
End Sub

' Every word (letters only) in s, as a case-insensitive lookup
Private Function KnownWords(ByVal s As String) As Object
    Dim d As Object, i As Long, wStart As Long, w As String
    Set d = CreateLookupDict()
    s = s & " "
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "[A-Za-z]" Then
            If wStart = 0 Then wStart = i
        ElseIf wStart > 0 Then
            w = Mid$(s, wStart, i - wStart)
            If Not DictExists(d, w) Then DictAdd d, w, True
            wStart = 0
        End If
    Next i
    Set KnownWords = d
End Function

' Keys, and text values, of a lookup dictionary as one string
Private Function DictWords(ByVal d As Object) As String
    Dim k As Variant, parts() As String, n As Long
    If d Is Nothing Then Exit Function
    If d.Count = 0 Then Exit Function
    ReDim parts(0 To 2 * d.Count)
    For Each k In d.Keys
        parts(n) = CStr(k)
        n = n + 1
        If VarType(d.Item(k)) = vbString Then
            parts(n) = d.Item(k)
            n = n + 1
        End If
    Next k
    DictWords = Join(parts, " ")
End Function

' ---- spelling ----------------------------------------------------------
' Corrects the clear misspellings and returns the others' positions (after the corrections)
' in redMarks, for the red underline
Private Sub ProofSpelling(ByVal doc As Object, ByVal known As Object, ByVal items As Collection, _
                          ByRef nFlag As Long, ByVal redMarks As Collection)
    Dim errs As Object, rng As Object, n As Long, i As Long, shift As Long
    Dim ss() As Long, se() As Long, words() As String, fixes() As String, sugText() As String
    Dim flagged() As Boolean, applied() As Boolean, cand As Variant

    Set errs = doc.SpellingErrors
    n = errs.Count
    If n = 0 Then Exit Sub
    If n > 500 Then n = 500
    ReDim ss(1 To n)
    ReDim se(1 To n)
    ReDim words(1 To n)
    ReDim fixes(1 To n)
    ReDim sugText(1 To n)
    ReDim flagged(1 To n)
    ReDim applied(1 To n)
    For i = 1 To n
        Set rng = errs(i)
        ss(i) = rng.Start
        se(i) = rng.End
        words(i) = rng.Text
    Next i

    ' decide first (reading only), in document order
    For i = 1 To n
        If SpellCandidate(words(i), known) And Not RepeatedWord(doc, ss(i), words(i)) Then
            cand = SpellingSuggestions(doc.Range(ss(i), se(i)))
            If Not IsInflection(words(i), cand) Then
                If PQ_PROOF_FIX Then fixes(i) = ClearSpellingFix(words(i), cand)
                If fixes(i) = "" And nFlag < PQ_PROOF_MAX Then
                    flagged(i) = True
                    nFlag = nFlag + 1
                    sugText(i) = SuggestionText(cand)
                End If
            End If
        End If
    Next i

    ' report in document order, with the sentence as it reads before any change
    For i = 1 To n
        If fixes(i) <> "" Then
            AddPQItem items, "Spelling corrected", words(i), fixes(i), PQV_PROOF, _
                      "corrected in the reviewed copy (Word's suggestion); correct the narrative the same way: " & _
                      ProofContext(doc, ss(i), se(i))
        ElseIf flagged(i) Then
            AddPQItem items, "Spelling", words(i), sugText(i), PQV_PROOF, _
                      "possible misspelling, not changed (wavy red underline): " & ProofContext(doc, ss(i), se(i))
        End If
    Next i

    ' change from the end, so earlier positions stay valid
    For i = n To 1 Step -1
        If fixes(i) <> "" Then
            Set rng = doc.Range(ss(i), se(i))
            If rng.Text = words(i) Then
                rng.Text = fixes(i)
                rng.Shading.BackgroundPatternColor = PQ_EYE
                applied(i) = True
            End If
        End If
    Next i
    For i = 1 To n
        If applied(i) Then
            shift = shift + Len(fixes(i)) - Len(words(i))
        ElseIf flagged(i) Then
            redMarks.Add Array(ss(i) + shift, se(i) + shift)
        End If
    Next i
End Sub

' Only plain lowercase words are judged: capitalised words are names or companies, words
' with digits or symbols are codes, and words from the transaction file are names
Private Function SpellCandidate(ByVal w As String, ByVal known As Object) As Boolean
    If Len(w) < 3 Then Exit Function
    If w Like "*[!A-Za-z]*" Then Exit Function
    If Left$(w, 1) <> LCase$(Left$(w, 1)) Then Exit Function
    If DictExists(known, w) Then Exit Function
    SpellCandidate = True
End Function

' Word reports the second of two equal words ("that that") as a spelling error; those the
' mechanical pass left alone are intended
Private Function RepeatedWord(ByVal doc As Object, ByVal s As Long, ByVal w As String) As Boolean
    Dim a As Long
    a = s - Len(w) - 1
    If a < doc.Content.Start Then Exit Function
    RepeatedWord = (StrComp(doc.Range(a, s).Text, w & " ", vbTextCompare) = 0)
End Function

' Word's suggestions for a range, at most five, as a 0-based array (empty array if none)
Private Function SpellingSuggestions(ByVal rng As Object) As Variant
    Dim sugs As Object, n As Long, i As Long, arr() As String
    SpellingSuggestions = Array()
    On Error Resume Next
    Set sugs = rng.GetSpellingSuggestions
    If sugs Is Nothing Then Exit Function
    n = sugs.Count
    If n > 5 Then n = 5
    If n <= 0 Then Exit Function
    ReDim arr(0 To n - 1)
    For i = 1 To n
        arr(i - 1) = sugs(i).Name
    Next i
    On Error GoTo 0
    SpellingSuggestions = arr
End Function

Private Function SuggestionText(ByVal cand As Variant) As String
    Dim i As Long, s As String
    For i = LBound(cand) To UBound(cand)
        If i - LBound(cand) >= 3 Then Exit For
        If s <> "" Then s = s & " / "
        s = s & CStr(cand(i))
    Next i
    If s = "" Then s = "(Word has no suggestion)"
    SuggestionText = s
End Function

' A plural or other form of a word Word does know (homestays, prefunded): not an error
Private Function IsInflection(ByVal w As String, ByVal cand As Variant) As Boolean
    Dim i As Long, c As String, sfx As Variant
    w = LCase$(w)
    For i = LBound(cand) To UBound(cand)
        c = LCase$(CStr(cand(i)))
        For Each sfx In Array("s", "es", "d", "ed", "ing", "ly", "er", "ers")
            If w = c & CStr(sfx) Then
                IsInflection = True
                Exit Function
            End If
        Next sfx
    Next i
End Function

' The suggestion to apply, or "": a single lowercase word, one typing slip away
' (a swapped pair or a doubled letter counts half), and strictly closer than every
' other suggestion. "funremains" -> "fun remains" or "parfum" -> "perfume" are never applied.
Private Function ClearSpellingFix(ByVal w As String, ByVal cand As Variant) As String
    Dim i As Long, c As String, d As Long, best As Long, bestWord As String, tie As Boolean
    If Len(w) < 4 Then Exit Function
    best = 99
    For i = LBound(cand) To UBound(cand)
        c = CStr(cand(i))
        If Len(c) >= 3 And Not (c Like "*[!a-z]*") Then
            d = TypoDistance(w, c)
            If d < best Then
                best = d
                bestWord = c
                tie = False
            ElseIf d = best Then
                tie = True
            End If
        End If
    Next i
    If best <= 2 And Not tie Then ClearSpellingFix = bestWord
End Function

' Edit distance in half-steps: an adjacent swap, or adding/dropping a letter next to the
' same letter, costs 1; any other insertion, deletion or substitution costs 2
Private Function TypoDistance(ByVal a As String, ByVal b As String) As Long
    Dim la As Long, lb As Long, i As Long, j As Long, c As Long, v As Long
    Dim d() As Long
    la = Len(a)
    lb = Len(b)
    ReDim d(0 To la, 0 To lb)
    For i = 1 To la
        d(i, 0) = d(i - 1, 0) + EditCost(a, i)
    Next i
    For j = 1 To lb
        d(0, j) = d(0, j - 1) + EditCost(b, j)
    Next j
    For i = 1 To la
        For j = 1 To lb
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then c = 0 Else c = 2
            v = d(i - 1, j - 1) + c
            If d(i - 1, j) + EditCost(a, i) < v Then v = d(i - 1, j) + EditCost(a, i)
            If d(i, j - 1) + EditCost(b, j) < v Then v = d(i, j - 1) + EditCost(b, j)
            If i > 1 And j > 1 Then
                If Mid$(a, i, 1) = Mid$(b, j - 1, 1) And Mid$(a, i - 1, 1) = Mid$(b, j, 1) Then
                    If d(i - 2, j - 2) + 1 < v Then v = d(i - 2, j - 2) + 1
                End If
            End If
            d(i, j) = v
        Next j
    Next i
    TypoDistance = d(la, lb)
End Function

' Cost of adding or dropping letter k of s: 1 when it doubles a neighbour, else 2
Private Function EditCost(ByVal s As String, ByVal k As Long) As Long
    EditCost = 2
    If k > 1 Then
        If Mid$(s, k - 1, 1) = Mid$(s, k, 1) Then EditCost = 1
    End If
    If k < Len(s) Then
        If Mid$(s, k + 1, 1) = Mid$(s, k, 1) Then EditCost = 1
    End If
End Function

' ---- grammar -----------------------------------------------------------
Private Sub ProofGrammar(ByVal doc As Object, ByVal items As Collection, ByRef nFlag As Long)
    Dim errs As Object, rng As Object, n As Long, i As Long, g As String

    Set errs = doc.GrammaticalErrors
    n = errs.Count
    For i = 1 To n
        If nFlag >= PQ_PROOF_MAX Then Exit For
        Set rng = errs(i)
        g = Trim$(Replace(Replace(rng.Text, vbCr, " "), Chr(11), " "))
        If g <> "" Then
            MarkProof rng, WD_COLOR_BLUE
            AddPQItem items, "Grammar", Left$(g, 120), "", PQV_PROOF, _
                      "Word flags this as a grammar issue and gives macros no correction; reword if needed " & _
                      "(wavy blue underline): " & ProofContext(doc, rng.Start, rng.End)
            nFlag = nFlag + 1
        End If
    Next i
End Sub

Private Sub MarkProof(ByVal rng As Object, ByVal colr As Long)
    On Error Resume Next
    rng.Font.Underline = WD_UNDERLINE_WAVY
    rng.Font.UnderlineColor = colr
    On Error GoTo 0
End Sub

' "...about 45 characters either side..." of a document range
Private Function ProofContext(ByVal doc As Object, ByVal s As Long, ByVal e As Long) As String
    Dim a As Long, b As Long, t As String
    a = s - 45
    If a < doc.Content.Start Then a = doc.Content.Start
    b = e + 45
    If b > doc.Content.End Then b = doc.Content.End
    t = doc.Range(a, b).Text
    t = Replace(Replace(Replace(t, vbCr, " / "), Chr(11), " "), Chr(7), " ")
    ProofContext = "..." & Trim$(t) & "..."
End Function

' ---- mechanical slips --------------------------------------------------
' A repeated word ("the the"), no space after a full stop ("behavior.The"), extra spaces
' between words, a space before a comma. Word flags these but gives a macro no fix.
' Paragraphs are handled from the last one up and each from its end, so earlier positions
' stay valid; the corrections are reported in document order.
Private Sub FixMechanical(ByVal doc As Object, ByVal items As Collection)
    Dim i As Long, k As Long, nPar As Long, pr As Object
    Dim perPara() As Collection, f As Variant

    nPar = doc.Paragraphs.Count
    If nPar = 0 Then Exit Sub
    ReDim perPara(1 To nPar)
    For i = nPar To 1 Step -1
        Set pr = doc.Paragraphs(i).Range
        Set perPara(i) = FixParagraph(doc, pr.Start, pr.Text)
    Next i
    For i = 1 To nPar
        For k = 1 To perPara(i).Count
            f = perPara(i)(k)
            AddPQItem items, "Grammar corrected", CStr(f(0)), CStr(f(1)), PQV_PROOF, CStr(f(2))
        Next k
    Next i
End Sub

' Finds the slips in one paragraph's text, applies them from the end, and returns
' Array(before, after, note) for each in paragraph order
Private Function FixParagraph(ByVal doc As Object, ByVal pStart As Long, ByVal t As String) As Collection
    Dim found As New Collection, done As New Collection
    Dim p As Long, n As Long, q As Long, k As Long, isStart As Boolean
    Dim w1 As String, w2 As String, gap As String, tok As String, ts As Long
    Dim f As Variant

    n = Len(t)
    p = 1
    Do While p <= n
        ' a word: letters from p to q-1, not part of a code, address or longer token
        isStart = False
        If Mid$(t, p, 1) Like "[A-Za-z]" Then
            If p = 1 Then
                isStart = True
            ElseIf Not (Mid$(t, p - 1, 1) Like "[A-Za-z0-9'.@/_-]") Then
                isStart = True
            End If
        End If
        If isStart Then
            q = p
            Do While q <= n
                If Not (Mid$(t, q, 1) Like "[A-Za-z]") Then Exit Do
                q = q + 1
            Loop
            w1 = Mid$(t, p, q - p)
            ' spaces after the word, then the next word
            k = q
            Do While k <= n
                If Mid$(t, k, 1) <> " " Then Exit Do
                k = k + 1
            Loop
            gap = Mid$(t, q, k - q)
            w2 = ""
            If gap <> "" And k <= n Then
                If Mid$(t, k, 1) Like "[A-Za-z]" Then w2 = NextWordAt(t, k)
            End If
            If w2 <> "" And w2 = w1 And w1 = LCase$(w1) And InStr(" had that is ", " " & w1 & " ") = 0 Then
                ' "the the" -> "the"
                If Not (Mid$(t, k + Len(w2), 1) Like "[A-Za-z0-9'-]") Then
                    found.Add Array("R", p, w1 & gap & w2, Len(w1), Len(gap) + Len(w2), "", _
                                    w1 & " " & w2, w1, "repeated word removed")
                End If
            ElseIf Len(gap) >= 2 And w2 <> "" Then
                ' "Ripoll  does" -> "Ripoll does"
                found.Add Array("S", p, w1 & gap & w2, Len(w1) + 1, Len(gap) - 1, "", _
                                w1 & gap & TokenAt(t, k), w1 & " " & TokenAt(t, k), "extra spaces removed")
            ElseIf gap = " " And Mid$(t, k, 1) = "," Then
                ' "funds ," -> "funds,"
                found.Add Array("S", p, w1 & " ,", Len(w1), 1, "", w1 & " ,", w1 & ",", "space before the comma removed")
            ElseIf gap = "" And Mid$(t, q, 1) = "." And Len(w1) >= 3 And w1 = LCase$(w1) Then
                ' "behavior.The" -> "behavior. The"
                w2 = NextWordAt(t, q + 1)
                If Len(w2) >= 2 Then
                    If Left$(w2, 1) Like "[A-Z]" And Mid$(w2, 2, 1) Like "[a-z]" _
                       And InStr(" com net org ai io co gov edu us uk ", " " & LCase$(w2) & " ") = 0 Then
                        found.Add Array("I", p, w1 & "." & w2, Len(w1) + 1, 0, " ", _
                                        w1 & "." & w2, w1 & ". " & w2, "space added after the full stop")
                    End If
                End If
            End If
            p = q
        Else
            p = p + 1
        End If
    Loop

    ' apply from the end of the paragraph
    For k = found.Count To 1 Step -1
        f = found(k)
        tok = CStr(f(2))
        ts = TruePos(doc, pStart + CLng(f(1)) - 1, Len(tok))
        If doc.Range(ts, ts + Len(tok)).Text = tok Then
            If CStr(f(0)) = "I" Then
                doc.Range(ts + CLng(f(3)), ts + CLng(f(3))).InsertAfter CStr(f(5))
            Else
                doc.Range(ts + CLng(f(3)), ts + CLng(f(3)) + CLng(f(4))).Delete
            End If
            If done.Count = 0 Then
                done.Add Array(f(6), f(7), f(8))
            Else
                done.Add Array(f(6), f(7), f(8)), Before:=1
            End If
        End If
    Next k
    Set FixParagraph = done
End Function

' The text from position k up to the next space (for the findings sheet)
Private Function TokenAt(ByVal t As String, ByVal k As Long) As String
    Dim q As Long
    q = InStr(k, t & " ", " ")
    TokenAt = Left$(Mid$(t, k, q - k), 30)
End Function

' The run of letters starting at position k
Private Function NextWordAt(ByVal t As String, ByVal k As Long) As String
    Dim q As Long
    q = k
    Do While q <= Len(t)
        If Not (Mid$(t, q, 1) Like "[A-Za-z]") Then Exit Do
        q = q + 1
    Loop
    NextWordAt = Mid$(t, k, q - k)
End Function

' "2 corrected, 3 to review" for the dashboard
Private Sub ProofCounts(ByVal items As Collection, ByRef nFix As Long, ByRef nRev As Long)
    Dim it As Variant
    nFix = 0
    nRev = 0
    For Each it In items
        If CStr(it(3)) = PQV_PROOF Then
            If Right$(CStr(it(0)), 9) = "corrected" Then nFix = nFix + 1 Else nRev = nRev + 1
        End If
    Next it
End Sub

Private Function ProofSummary(ByVal items As Collection) As String
    Dim nFix As Long, nRev As Long
    If Not PQ_PROOF Then Exit Function
    ProofCounts items, nFix, nRev
    ProofSummary = nFix & " corrected, " & nRev & " to review"
End Function
