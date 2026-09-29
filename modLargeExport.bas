Attribute VB_Name = "modLargeExport"
Option Explicit

' ==========================================================
' LARGE-FILE EXPORT (POWER QUERY) - a separate copy of Export Trx File
' ==========================================================
' For cases with hundreds of thousands of transaction rows, which the
' normal Export Trx File in Module9 can't handle. Same three export modes,
' same output files and names, same pivots. Run it from
' Developer > Macros > Consolidated_AML_Workflow_Large; the Export Trx
' File button still runs the normal Module9 version.
'
' Module9 opens every source file in Excel, pastes it under the rows
' already combined, then copies that sheet around and deletes the rows
' each export doesn't want. At 6 lakh rows that is several full copies of
' the data in memory, and the paste itself fails ("the Copy area and
' paste area aren't the same size") when a file's used range runs far
' below its data, or when the files add up to more rows than a sheet
' holds.
'
' Here Power Query reads the files instead. It combines the folder, does
' the date / amount / Counterparty cleanup and filters each output sheet,
' and each sheet is loaded straight into its export workbook. The queries
' are deleted again before saving, so the exports hold plain data as
' Module9's do. VBA still does the export picker, the date windows, the
' pivots, saving, ConsolidatedData and the tracker.
'
' Differences from Module9's output:
'   - Counterparty is a value, not a formula.
'   - Numbers a source file holds as text stay text (Module9's value
'     rewrite turned them into numbers). Amounts are still converted.
'   - Files are matched up by column NAME when combined; Module9 stacked
'     them by column position.
'   - A file's "Transaction ID" header must be in its first 100 rows.
'   - The Pivot Analysis export, which is saved into the same \Pivot
'     folder it reads from, is not read back in as data on the next run.
'   - EN's Raw Transactions carries on onto "Raw Transactions (2)",
'     "(3)"... when it has more rows than one sheet holds.
'   - Legacy leaves out Raw Transactions (the files themselves are in the
'     Transaction Files folder), and CP Selection and DeDupe carry on onto
'     "(2)", "(3)"... sheets. Its four pivots are built on two totals
'     sheets Power Query adds up (Scenario Totals, Daily Totals), so they
'     work however many sheets the rows are spread over; the pivots add
'     the totals back up, so every figure matches a pivot on the rows.
'   - When Legacy's DeDupe is too big for ConsolidatedData, ConsolidatedData
'     gets only the rows that carry an alert, and Sheet7's narrative
'     figures are worked out by Power Query over every transaction - see
'     PointSheet7AtLargeStats.
'   - Any other sheet that would need more rows than Excel allows stops
'     the export with a message instead of being cut short.
'   - Legacy's duplicate check treats a Transaction ID held as a number in
'     one file and as text in another as the same ID, and ignores case -
'     as RemoveDuplicates did after Module9 turned both into numbers.
' ==========================================================

' TidyDataSheet: sheets up to this many rows get the full column + wrapped
' row fit; bigger ones fit columns to the first TIDY_SAMPLE_ROWS rows only.
Private Const TIDY_FULL_FIT_ROWS As Long = 50000
Private Const TIDY_SAMPLE_ROWS As Long = 2000

' Rows per Raw Transactions sheet when Raw has to be spread over several
' (a sheet holds 1,048,575 below its header).
Private Const RAW_ROWS_PER_SHEET As Long = 1000000

' The step the export is on - shown on the status bar while it runs and
' named in the error message if it fails.
Private m_stage As String

' Export workbooks created this run and not yet saved.
Private m_unsaved As Collection

' Very hidden sheet in this workbook holding Sheet7's narrative figures for
' a case too big for ConsolidatedData (B1 = that case's ECM ID).
Private Const STATS_SHEET As String = "_LargeCaseStats"

Sub Consolidated_AML_Workflow_Large()

' ==========================================
' LARGE-FILE EXPORT (POWER QUERY)
' ==========================================
Application.EnableCancelKey = xlErrorHandler
On Error GoTo CancelHandler
m_stage = "starting the export"
Set m_unsaved = New Collection

' Safe default so CancelHandler can always restore Calculation even
' if an error fires before the real capture below ever runs.
Dim origCalc As XlCalculation
origCalc = xlCalculationAutomatic

Dim wsHome As Worksheet, wsRealCD As Worksheet, ws As Worksheet
Dim newWb As Workbook, lbWb As Workbook
Dim desktopPath As String, saveFolderPath As String, folderPath As String, slash As String
Dim excelFileName As String, finalSavePath As String
Dim ecmID As String, AlertID As String
Dim FSO As Object, objFolder As Object, objFile As Object
Dim fileFound As Boolean
Dim wsPivot As Worksheet, ptCache As PivotCache, pt As PivotTable, ptRange As Range
Dim lastRowCP As Long, lastColCP As Long
Dim wsExport As Worksheet, wsDeDupe As Worksheet

' Power Query state: the combine query every export workbook starts from,
' the cleanup query built on it, and the real header names they use.
Dim excluded As Collection, baseM As String, cleanM As String, hdr As Variant
Dim colDate As String, colAmt As String, colDrCr As String, colBen As String
Dim colOrig As String, colFlag As String, colTrans As String, colAlert As String, colCp As String

' --- THE ULTIMATE PATH FIX ---
slash = Application.PathSeparator

' Break sheet grouping if active
On Error Resume Next
If Not ActiveWindow Is Nothing Then
    If ActiveWindow.SelectedSheets.count > 1 Then
        ActiveSheet.Select
    End If
End If
On Error GoTo CancelHandler

' ==========================================
' 1. SETUP (same checks as Module9)
' ==========================================
ThisWorkbook.Unprotect Password:="p7ss"
On Error Resume Next
ThisWorkbook.Sheets("Sheet1").Unprotect Password:="p7ss"
ThisWorkbook.Sheets("ConsolidatedData").Unprotect Password:="p7ss"
On Error GoTo CancelHandler

' ThisWorkbook rather than Module9's ActiveWorkbook: this is run from the
' Macros dialog, where the active workbook may well be the last export.
Set wsHome = ThisWorkbook.Sheets("Sheet1")
ecmID = Trim(wsHome.Range("J9").Value)
AlertID = Trim(wsHome.Range("J10").Value)
If AlertID = "" Then AlertID = "ALERT"

If ecmID = "" Then
    MsgBox "Action Denied: ECM ID is missing in J9.", vbCritical, "Missing ID"
    Exit Sub
End If

' ==========================================
' 1b. EXPORT FORMAT PICKER - frmExportMode
' ==========================================
Dim exportMode As String
frmExportMode.Show vbModal

If frmExportMode.userCancelled Then
    Unload frmExportMode
    Exit Sub
End If
exportMode = frmExportMode.SelectedMode   ' "LEGACY" / "EN" / "PIVOT"
Unload frmExportMode

Dim sourceFolderName As String
If exportMode = "PIVOT" Then
    sourceFolderName = "Pivot"
Else
    sourceFolderName = "Transaction Files"
End If

desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
saveFolderPath = desktopPath & slash & ecmID
folderPath = saveFolderPath & slash & sourceFolderName

Set FSO = CreateObject("Scripting.FileSystemObject")

' CHECK IF FOLDER EXISTS
If Not FSO.FolderExists(folderPath) Then
    MsgBox "Folder structure missing!" & vbCrLf & "Please click the START button to generate your '" & ecmID & "" & sourceFolderName & "' folder first.", vbCritical, "Folder Not Found"
    Exit Sub
End If

' CHECK IF FILES EXIST
Set objFolder = FSO.GetFolder(folderPath)
fileFound = False

For Each objFile In objFolder.Files
    If (InStr(1, objFile.Name, ".xls", vbTextCompare) > 0) And (Left(objFile.Name, 2) <> "~$") And (objFile.Name <> ThisWorkbook.Name) Then
        fileFound = True
        Exit For
    End If
Next objFile

If Not fileFound Then
    MsgBox "No Excel files found in the target folder!", vbExclamation, "Folder is Empty"
    Exit Sub
End If

' Files Power Query must not read as transaction data: this workbook, and
' in Pivot mode the Pivot Analysis export, which is saved into the same
' \Pivot folder it reads from and would otherwise come back in as extra
' rows on the next run.
Dim pivotFileName As String
pivotFileName = ecmID & "_" & AlertID & "_Pivot Analysis.xlsx"
Set excluded = New Collection
excluded.Add ThisWorkbook.Name
If exportMode = "PIVOT" Then
    excluded.Add pivotFileName
    excluded.Add PreviousExportPath(pivotFileName)
End If
baseM = BuildBaseM(folderPath & slash, excluded)

origCalc = Application.Calculation
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' ==========================================
' 1c. PIVOT ANALYSIS EXPORT (own source folder: \Pivot)
' ==========================================
If exportMode = "PIVOT" Then
    Dim newWbPiv As Workbook, pSh As Long, pivotSavePath As String

    SetStage "reading the column headers"
    Set newWbPiv = NewOutputWorkbook()
    newWbPiv.Sheets(1).Name = "Pivot Data"
    newWbPiv.Queries.Add Name:="TrxBase", Formula:=baseM
    hdr = ReadHeaderNames(newWbPiv)
    If IsEmpty(hdr) Then
        MsgBox NoHeaderMessage(folderPath), vbCritical, "No Transaction Data"
        GoTo CancelHandler
    End If
    ResolveColumns hdr, colDate, colAmt, colDrCr, colBen, colOrig, colFlag, colTrans, colAlert, colCp
    newWbPiv.Queries.Add Name:="TrxClean", Formula:=BuildCleanM(colDate, colAmt, colDrCr, colBen, colOrig, colCp)

    SetStage "loading the Pivot data"
    LoadQueryToSheet newWbPiv.Sheets("Pivot Data"), "TrxClean"
    RemoveAllQueries newWbPiv
    FormatDataColumns newWbPiv.Sheets("Pivot Data"), colDate, "m/d/yyyy", colAmt

    SetStage "building the Pivot Analysis pivots"
    BuildEnPivots newWbPiv, "Pivot Data", "Pivot", "Pivot Data"

    For pSh = newWbPiv.Sheets.count To 1 Step -1
        Select Case newWbPiv.Sheets(pSh).Name
            Case "Pivot Data", "Pivot"
                ' keep
            Case Else
                SafeDeleteSheet newWbPiv, newWbPiv.Sheets(pSh).Name
        End Select
    Next pSh

    SetStage "formatting the Pivot Data sheet"
    TidyDataSheet newWbPiv.Sheets("Pivot Data")

    pivotSavePath = folderPath & slash & pivotFileName
    SetStage "saving the Pivot Analysis file"
    CloseIfAlreadyOpen pivotSavePath
    MoveExistingExportAside pivotSavePath
    SetSheetZoom85 newWbPiv, Array("Pivot Data", "Pivot")
    Application.DisplayAlerts = False
    newWbPiv.SaveAs fileName:=pivotSavePath, FileFormat:=51
    MarkSaved newWbPiv
    DiscardPreviousExport pivotSavePath

    newWbPiv.Sheets("Pivot Data").Activate
    FinishRun origCalc

    MsgBox "Pivot Analysis export complete!" & vbCrLf & _
        "ConsolidatedData was not touched." & vbCrLf & _
        "Saved to:" & vbCrLf & pivotSavePath, vbInformation, "Success"
    Exit Sub
End If

' Initialize the REAL ConsolidatedData sheet
On Error Resume Next
Set wsRealCD = ThisWorkbook.Sheets("ConsolidatedData")
On Error GoTo CancelHandler

If wsRealCD Is Nothing Then
    Set wsRealCD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsRealCD.Name = "ConsolidatedData"
End If

' ==========================================
' 2-3. COMBINE + CLEANUP (Power Query)
' ==========================================
' The queries live in the export workbook they fill, and are deleted
' again before it is saved. Reading the header row first means missing
' columns get Module9's own messages, and the cleanup and filters can be
' written against the real column names.
SetStage "reading the column headers"
Set newWb = NewOutputWorkbook()
newWb.Queries.Add Name:="TrxBase", Formula:=baseM
hdr = ReadHeaderNames(newWb)
If IsEmpty(hdr) Then
    MsgBox NoHeaderMessage(folderPath), vbCritical, "No Transaction Data"
    GoTo CancelHandler
End If
ResolveColumns hdr, colDate, colAmt, colDrCr, colBen, colOrig, colFlag, colTrans, colAlert, colCp
cleanM = BuildCleanM(colDate, colAmt, colDrCr, colBen, colOrig, colCp)
newWb.Queries.Add Name:="TrxClean", Formula:=cleanM

' ==========================================
' 4-EN. EN NETWORK EXPORT
' ==========================================
If exportMode = "EN" Then
    Dim wsLB As Worksheet, wsRawEN As Worksheet, wsAlertedEN As Worksheet, wsNonEN As Worksheet
    Dim firstAlerted As Date, lastAlerted As Date
    Dim lbStart As Date, lbEnd As Date, winStartEN As Date, winEndEN As Date
    Dim lbSavedPath As String, nonCount As Long, rEN As Long

    If Len(colFlag) = 0 Then
        MsgBox "EN Network export needs an 'Is Alerted Transaction?' column (exact name), but it wasn't found in the data.", vbCritical, "Column Not Found"
        GoTo CancelHandler
    End If

    Set wsRawEN = newWb.Sheets(1)
    wsRawEN.Name = "Raw Transactions"
    Set wsAlertedEN = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.count))
    wsAlertedEN.Name = "Alerted Transaction"
    Set wsNonEN = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.count))
    wsNonEN.Name = "Non Alerted Transaction"

    ' The alerted rows are loaded first: their dates set both windows below.
    SetStage "loading the alerted transactions"
    newWb.Queries.Add Name:="TrxAlerted", Formula:=BuildFilterM("TrxClean", colFlag, "Yes", "", False, 0, 0)
    LoadQueryToSheet wsAlertedEN, "TrxAlerted"
    ' Date format first: a date only reads back as a Date once its cell is
    ' formatted as one.
    FormatDataColumns wsAlertedEN, colDate, "m/d/yyyy", colAmt

    If Not AlertedDateSpan(wsAlertedEN, colDate, firstAlerted, lastAlerted) Then
        MsgBox "No dated 'Yes' alerted transactions were found, so the EN Network export can't be built.", vbCritical, "No Alerted Rows"
        GoTo CancelHandler
    End If

    ' Same windows as Module9 (its comments explain both):
    ' Lookback - first day of the first alerted month one year back, to the
    ' last day of the last alerted month.
    lbStart = DateSerial(Year(firstAlerted) - 1, Month(firstAlerted), 1)
    lbEnd = DateSerial(Year(lastAlerted), Month(lastAlerted) + 1, 0)
    ' Non Alerted - the alerted month(s) only.
    winStartEN = DateSerial(Year(firstAlerted), Month(firstAlerted), 1)
    winEndEN = DateSerial(Year(lastAlerted), Month(lastAlerted) + 1, 0)

    ' ---------------------------------------------------------------
    ' Lookback Transactions file
    ' ---------------------------------------------------------------
    SetStage "loading the Lookback transactions"
    Set lbWb = NewOutputWorkbook()
    Set wsLB = lbWb.Sheets(1)
    wsLB.Name = "Lookback Transactions"
    lbWb.Queries.Add Name:="TrxBase", Formula:=baseM
    lbWb.Queries.Add Name:="TrxClean", Formula:=cleanM
    lbWb.Queries.Add Name:="TrxLookback", Formula:=BuildFilterM("TrxClean", "", "", colDate, True, lbStart, lbEnd)
    LoadQueryToSheet wsLB, "TrxLookback"
    RemoveAllQueries lbWb
    FormatDataColumns wsLB, colDate, "dddd, mmmm d, yyyy", colAmt

    SetStage "formatting the Lookback sheet"
    TidyDataSheet wsLB

    SetStage "building the Lookback pivots"
    BuildEnPivots lbWb, "Lookback Transactions", "Pivot", "Lookback Transactions"

    For rEN = lbWb.Sheets.count To 1 Step -1
        Select Case lbWb.Sheets(rEN).Name
            Case "Lookback Transactions", "Pivot"
                ' keep
            Case Else
                SafeDeleteSheet lbWb, lbWb.Sheets(rEN).Name
        End Select
    Next rEN

    excelFileName = ecmID & "_" & AlertID & "_Lookback Transactions (" & _
        Format$(lbStart, "mm.dd.yyyy") & " to " & Format$(lbEnd, "mm.dd.yyyy") & ").xlsx"
    finalSavePath = saveFolderPath & slash & excelFileName
    SetStage "saving the Lookback file"
    CloseIfAlreadyOpen finalSavePath
    MoveExistingExportAside finalSavePath
    SetSheetZoom85 lbWb, Array("Lookback Transactions", "Pivot")
    Application.DisplayAlerts = False
    lbWb.SaveAs fileName:=finalSavePath, FileFormat:=51
    MarkSaved lbWb
    DiscardPreviousExport finalSavePath
    lbSavedPath = finalSavePath

    ' ---------------------------------------------------------------
    ' Alerted / Non-Alerted Trx File
    ' ---------------------------------------------------------------
    SetStage "loading the Raw Transactions"
    LoadQueryAcrossSheets newWb, wsRawEN, "TrxClean"

    SetStage "loading the non-alerted transactions"
    newWb.Queries.Add Name:="TrxNonAlerted", Formula:=BuildFilterM("TrxClean", colFlag, "No", colDate, True, winStartEN, winEndEN)
    nonCount = LoadQueryToSheet(wsNonEN, "TrxNonAlerted")
    RemoveAllQueries newWb

    If nonCount = 0 Then
        wsNonEN.Cells.Clear
        wsNonEN.Range("A1").Value = "There were 0 non-alerted transactions during the alerted month(s) " & _
            Format$(winStartEN, "mm/dd/yyyy") & " to " & Format$(winEndEN, "mm/dd/yyyy") & "."
    End If

    SetStage "formatting the Combined sheets"
    For Each ws In newWb.Worksheets
        If IsSheetPart(ws.Name, "Raw Transactions") Or ws.Name = "Alerted Transaction" Or ws.Name = "Non Alerted Transaction" Then
            FormatDataColumns ws, colDate, "m/d/yyyy", colAmt
            TidyDataSheet ws
        End If
    Next ws

    SetStage "building the Alerted pivots"
    BuildEnPivots newWb, "Alerted Transaction", "Alerted Transaction Pivot", "Alerted Transaction"

    For rEN = newWb.Sheets.count To 1 Step -1
        Select Case newWb.Sheets(rEN).Name
            Case "Raw Transactions", "Alerted Transaction", "Alerted Transaction Pivot", "Non Alerted Transaction"
                ' keep
            Case Else
                If Not IsSheetPart(newWb.Sheets(rEN).Name, "Raw Transactions") Then SafeDeleteSheet newWb, newWb.Sheets(rEN).Name
        End Select
    Next rEN

    newWb.Sheets("Non Alerted Transaction").Move After:=newWb.Sheets(newWb.Sheets.count)

    excelFileName = ecmID & "_" & AlertID & "_Combined Alerted & Non Alerted Transactions.xlsx"
    finalSavePath = saveFolderPath & slash & excelFileName
    SetStage "saving the Combined file"
    CloseIfAlreadyOpen finalSavePath
    MoveExistingExportAside finalSavePath
    ZoomAllSheets newWb
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
    MarkSaved newWb
    DiscardPreviousExport finalSavePath

    SetStage "updating ConsolidatedData"
    wsRealCD.Cells.Clear
    newWb.Sheets("Alerted Transaction").UsedRange.Copy Destination:=wsRealCD.Range("A1")
    ClearLargeCaseStats
    On Error Resume Next
    Module3.RefreshRuleNameTag
    On Error GoTo CancelHandler

    newWb.Sheets("Raw Transactions").Activate
    FinishRun origCalc

    MsgBox "EN Network export complete! Both files were generated:" & vbCrLf & vbCrLf & _
        "Lookback Transactions:" & vbCrLf & lbSavedPath & vbCrLf & vbCrLf & _
        "Alerted / Non-Alerted Transactions:" & vbCrLf & finalSavePath, vbInformation, "Success"
    Exit Sub
End If

' ==========================================
' 4. TWO-LAYER EXPORT & DEDUPE (LEGACY)
' ==========================================
' CP Selection is the cleaned data less duplicate Transaction ID + Alert
' Information pairs, and DeDupe that less duplicate Transaction IDs and
' rows with no date - Module9's two RemoveDuplicates passes. Each carries
' on onto "(2)", "(3)"... sheets past RAW_ROWS_PER_SHEET rows.
'
' Raw Transactions (every row of every file, duplicates included) is left
' out: the files themselves are in the Transaction Files folder, and with
' several overlapping files it would be the biggest part of the export.
'
' The four pivots are built on two small totals sheets Power Query adds up
' - Scenario Totals (per alert / Dr Cr / counterparty, from CP Selection)
' and Daily Totals (per day / Dr Cr, from DeDupe) - not on the data sheets,
' which can be spread over several sheets. A pivot adds the totals back
' up, so every figure is the one a pivot on the rows would show.
Dim colAcct As String, cpRows As Long, ddRows As Long, alertRows As Long, shIdx As Long
Dim wsScnTot As Worksheet, wsDayTot As Worksheet, wsAlertTmp As Worksheet
Dim bigCase As Boolean, stats As Variant, unmatched As String, doneMsg As String

colAcct = FindHeaderName(hdr, "Account No", False)

If Len(colTrans) > 0 And Len(colAlert) > 0 Then
    newWb.Queries.Add Name:="TrxCP", Formula:=BuildDedupeM("TrxClean", Array(colTrans, colAlert), "")
Else
    newWb.Queries.Add Name:="TrxCP", Formula:=BuildDedupeM("TrxClean", Empty, "")
End If
If Len(colTrans) > 0 Then
    newWb.Queries.Add Name:="TrxDeDupe", Formula:=BuildDedupeM("TrxCP", Array(colTrans), colDate)
Else
    newWb.Queries.Add Name:="TrxDeDupe", Formula:=BuildDedupeM("TrxCP", Empty, colDate)
End If
newWb.Queries.Add Name:="TrxScenarioTotals", Formula:=BuildTotalsM("TrxCP", _
    Array(colAlert, colDrCr, colCp), Array("Alert Information", "Dr Cr", "Counterparty"), colAmt, "")
newWb.Queries.Add Name:="TrxDailyTotals", Formula:=BuildTotalsM("TrxDeDupe", _
    Array(colDate, colDrCr), Array("Transaction Date", "Dr Cr"), colAmt, colDate)

Set wsExport = newWb.Sheets(1)
wsExport.Name = "CP Selection"
Set wsDeDupe = newWb.Sheets.Add(After:=wsExport)
wsDeDupe.Name = "DeDupe"

SetStage "loading the CP Selection rows"
cpRows = LoadQueryAcrossSheets(newWb, wsExport, "TrxCP")
SetStage "loading the DeDupe rows"
ddRows = LoadQueryAcrossSheets(newWb, wsDeDupe, "TrxDeDupe")

SetStage "adding up the pivot totals"
Set wsScnTot = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.count))
wsScnTot.Name = "Scenario Totals"
LoadQueryToSheet wsScnTot, "TrxScenarioTotals"
Set wsDayTot = newWb.Sheets.Add(After:=wsScnTot)
wsDayTot.Name = "Daily Totals"
LoadQueryToSheet wsDayTot, "TrxDailyTotals"

' Too many transactions for one sheet means too many for ConsolidatedData.
' It gets the rows that carry an alert instead - every rule name Generate
' Narrative looks up is on those - and Sheet7's figures are worked out
' here over all of them.
bigCase = (ddRows > RAW_ROWS_PER_SHEET)
If bigCase Then
    SetStage "working out the narrative figures"
    newWb.Queries.Add Name:="TrxStats", Formula:=BuildStatsM(colAmt, colDate, colDrCr, colAcct)
    stats = ReadQueryRow(newWb, "TrxStats")
    If Not IsArray(stats) Then
        Err.Raise vbObjectError + 1005, "modLargeExport", "The narrative figures query returned nothing."
    End If
    If Len(colAlert) > 0 Then
        SetStage "loading the alerted rows for ConsolidatedData"
        newWb.Queries.Add Name:="TrxAlertRows", Formula:=BuildDedupeM("TrxDeDupe", Empty, colAlert)
        Set wsAlertTmp = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.count))
        wsAlertTmp.Name = "_AlertRows"
        alertRows = LoadQueryToSheet(wsAlertTmp, "TrxAlertRows")
    End If
End If
RemoveAllQueries newWb

' Backwards, because sheets are deleted along the way.
SetStage "formatting the Legacy sheets"
For shIdx = newWb.Worksheets.count To 1 Step -1
    Set ws = newWb.Worksheets(shIdx)
    If IsSheetPart(ws.Name, "CP Selection") Then
        FormatDataColumns ws, colDate, "dddd, mmmm d, yyyy", colAmt
        TidyDataSheet ws
    ElseIf IsSheetPart(ws.Name, "DeDupe") Then
        FormatDataColumns ws, colDate, "m/d/yyyy", colAmt
        TidyDataSheet ws
    ElseIf ws.Name = "_AlertRows" Then
        FormatDataColumns ws, colDate, "m/d/yyyy", colAmt
    ElseIf ws.Name = "Scenario Totals" Or ws.Name = "Daily Totals" Then
        FormatDataColumns ws, "Transaction Date", "m/d/yyyy", "Transaction Amount"
        TidyDataSheet ws
    Else
        SafeDeleteSheet newWb, ws.Name
    End If
Next shIdx

' ==========================================
' 4.5 PIVOT TABLES (from the totals sheets)
' ==========================================
SetStage "building the Legacy pivots"
Set wsPivot = newWb.Sheets.Add(Before:=newWb.Sheets(1))
wsPivot.Name = "Pivot"
BuildTotalsPivots newWb, wsPivot, wsScnTot, wsDayTot

' A big case's ConsolidatedData comes from the temporary alerted-rows
' sheet, which must not be saved into the export, so it is filled here,
' before the save.
If bigCase Then
    SetStage "updating ConsolidatedData"
    wsRealCD.Cells.Clear
    If Not wsAlertTmp Is Nothing Then
        wsAlertTmp.UsedRange.Copy Destination:=wsRealCD.Range("A1")
        Application.DisplayAlerts = False
        wsAlertTmp.Delete
    End If
    WriteLargeCaseStats stats, wsHome.Range("J9").Value
    unmatched = PointSheet7AtLargeStats()
End If

' ==========================================
' 5. FINALIZE MASTER TAB & SAVE
' ==========================================
Dim fileTag As String
fileTag = "Alerted"
excelFileName = ecmID & "_" & AlertID & "_Combined_" & fileTag & "_Transaction.xlsx"

finalSavePath = saveFolderPath & slash & excelFileName
SetStage "saving the Legacy file"
CloseIfAlreadyOpen finalSavePath
MoveExistingExportAside finalSavePath
ZoomAllSheets newWb

Application.DisplayAlerts = False
newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
MarkSaved newWb
DiscardPreviousExport finalSavePath

If Not bigCase Then
    SetStage "updating ConsolidatedData"
    wsRealCD.Cells.Clear
    newWb.Sheets("DeDupe").UsedRange.Copy Destination:=wsRealCD.Range("A1")
    ClearLargeCaseStats
End If
TidyDataSheet wsRealCD

On Error Resume Next
Module3.RefreshRuleNameTag
On Error GoTo CancelHandler

newWb.Sheets("Pivot").Activate
FinishRun origCalc

doneMsg = "Workflow Complete!" & vbCrLf & _
    "Exported file inside the folder exactly to: " & vbCrLf & finalSavePath
If bigCase Then
    doneMsg = doneMsg & vbCrLf & vbCrLf & _
        "There are " & Format$(ddRows, "#,##0") & " de-duplicated transactions - more than one sheet " & _
        "holds - so DeDupe and CP Selection carry on onto further sheets." & vbCrLf & vbCrLf & _
        "ConsolidatedData holds only the " & Format$(alertRows, "#,##0") & " rows that carry an alert. " & _
        "Sheet7's narrative figures (count, totals, date range, CR/DR, account numbers) were worked out " & _
        "from all " & Format$(ddRows, "#,##0") & " transactions, and apply while Sheet1 has this ECM ID."
    If Len(unmatched) > 0 Then
        doneMsg = doneMsg & vbCrLf & vbCrLf & "Check Sheet7 " & unmatched & ": these read ConsolidatedData " & _
            "but weren't recognised, so they only see the alerted rows."
    End If
End If
MsgBox doneMsg, vbInformation, "Success"
Exit Sub

CancelHandler:
Dim savedErrNum As Long, savedErrDesc As String
savedErrNum = Err.Number
savedErrDesc = Err.Description
' Leave error-handling mode, so the cleanup below can trap its own errors.
On Error GoTo -1

On Error Resume Next
' An export workbook that never got saved is closed unsaved - otherwise a
' failed run leaves hundreds of thousands of rows open in a "BookN" window.
CloseUnsavedExports
Application.Calculation = origCalc
Application.EnableCancelKey = xlInterrupt
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = False

ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False

Application.EnableEvents = True
On Error GoTo 0

If savedErrNum = 18 Then
    MsgBox "Process Safely Cancelled.", vbInformation, "Aborted"
ElseIf savedErrNum <> 0 Then
    MsgBox "An unexpected error occurred while " & m_stage & ":" & vbCrLf & vbCrLf & _
        "Error " & savedErrNum & ": " & savedErrDesc, vbCritical, "Large Export Error"
End If
End Sub

' ==========================================================
' SetStage - records the current step and shows it on the status bar
' ==========================================================
Private Sub SetStage(ByVal stage As String)
    m_stage = stage
    Application.StatusBar = "Large export: " & stage & "..."
End Sub

' ==========================================================
' PivotSourceAddress - a pivot cache's source as text, not a Range
' ==========================================================
' PivotCaches.Create is known to fail with "Type mismatch" when SourceData
' is a Range object of more than 65,536 rows. The same range given as an
' address string works at any size: 'Sheet name'!R1C1:R600001C38.
' ==========================================================
Private Function PivotSourceAddress(ByVal rng As Range) As String
    PivotSourceAddress = "'" & Replace(rng.Worksheet.Name, "'", "''") & "'!" & _
        rng.Address(ReferenceStyle:=xlR1C1)
End Function

' ==========================================================
' SafeDeleteSheet - 100% immune to Subscript out of range (Err 9)
' and Workbook Protection / VeryHidden deletion crashes (Err 1004)
' ==========================================================
Private Sub SafeDeleteSheet(ByVal wb As Workbook, ByVal sheetName As String)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub

    wb.Unprotect Password:="p7ss"
    wb.Unprotect

    Dim wsTarget As Worksheet
    Set wsTarget = Nothing
    Set wsTarget = wb.Sheets(sheetName)

    If Not wsTarget Is Nothing Then
        Application.DisplayAlerts = False
        wsTarget.Visible = xlSheetVisible
        wsTarget.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub

' Scans a sheet's UsedRange for date-shaped values and forces mm/dd/yyyy
Private Sub BulletproofDateFormat(ByVal ws As Worksheet)
On Error Resume Next

Dim rUsed As Range
Set rUsed = ws.UsedRange

If rUsed.Cells.count = 1 Then
    If IsDate(rUsed.Value) And Not IsEmpty(rUsed.Value) Then
        If Year(CDate(rUsed.Value)) > 1950 Then rUsed.NumberFormat = "mm/dd/yyyy"
    End If
    Exit Sub
End If

Dim arr As Variant, r As Long, c As Long
Dim hitRange As Range
arr = rUsed.Value

For r = 1 To UBound(arr, 1)
    For c = 1 To UBound(arr, 2)
        If IsDate(arr(r, c)) And Not IsEmpty(arr(r, c)) Then
            If Year(CDate(arr(r, c))) > 1950 Then
                If hitRange Is Nothing Then
                    Set hitRange = rUsed.Cells(r, c)
                Else
                    Set hitRange = Union(hitRange, rUsed.Cells(r, c))
                End If
            End If
        End If
    Next c
Next r

If Not hitRange Is Nothing Then hitRange.NumberFormat = "mm/dd/yyyy"
End Sub

' Highlights the Sum-column cell for every CR/DR row in the pivot's RowRange
Private Sub HighlightDrCrRows(ByVal pt As PivotTable, ByVal ws As Worksheet)
On Error Resume Next

Dim sumColIndex As Long, baseRow As Long
sumColIndex = pt.DataBodyRange.Columns(2).Column
baseRow = pt.RowRange.row

Dim rowArr As Variant, rIdx As Long, cellVal As String
Dim hitRange As Range
rowArr = pt.RowRange.Value

If IsArray(rowArr) Then
    For rIdx = 1 To UBound(rowArr, 1)
        cellVal = Trim(UCase(CStr(rowArr(rIdx, 1))))
        If cellVal = "CR" Or cellVal = "DR" Then
            If hitRange Is Nothing Then
                Set hitRange = ws.Cells(baseRow + rIdx - 1, sumColIndex)
            Else
                Set hitRange = Union(hitRange, ws.Cells(baseRow + rIdx - 1, sumColIndex))
            End If
        End If
    Next rIdx
Else
    cellVal = Trim(UCase(CStr(rowArr)))
    If cellVal = "CR" Or cellVal = "DR" Then
        Set hitRange = ws.Cells(baseRow, sumColIndex)
    End If
End If

If Not hitRange Is Nothing Then
    hitRange.Interior.Color = RGB(255, 199, 206)   ' Light Red Fill
    hitRange.Font.Color = RGB(156, 0, 6)           ' Dark Red Text
End If
End Sub

' ==========================================================
' BuildEnPivots - builds the 4 "Legacy" pivots on a fresh pivot sheet
' ==========================================================
Private Sub BuildEnPivots(ByVal wb As Workbook, ByVal dataSheet As String, _
    ByVal pivotSheet As String, ByVal insertAfter As String)
    Dim wsData As Worksheet, wsPv As Worksheet, wsPvScratch As Worksheet
    Dim dCol As Long, lastRow As Long, lastCol As Long, ddLast As Long
    Dim rngScn As Range, cacheScn As PivotCache, rngTmp As Range, cacheTmp As PivotCache
    Dim ptx As PivotTable, ptE As PivotTable, ptD As PivotTable

    On Error Resume Next
    Set wsData = wb.Sheets(dataSheet)
    On Error GoTo 0
    If wsData Is Nothing Then Exit Sub

    dCol = 0
    On Error Resume Next
    dCol = wsData.Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
    On Error GoTo 0

    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    lastCol = wsData.Cells(1, wsData.Columns.count).End(xlToLeft).Column
    If lastRow <= 1 Then Exit Sub

    Set wsPv = wb.Sheets.Add(After:=wb.Sheets(insertAfter))
    wsPv.Name = pivotSheet

    ' PIVOT 1: SCENARIO
    Set rngScn = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
    Set cacheScn = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PivotSourceAddress(rngScn))
    Set ptx = cacheScn.CreatePivotTable(TableDestination:=wsPv.Range("A3"), TableName:="ScenarioPivot")
    On Error Resume Next
    With ptx
        .TableStyle2 = "PivotStyleLight16"
        With .PivotFields("Alert Information"): .Orientation = xlRowField: .Position = 1: End With
        With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 2: End With
        With .PivotFields("Counterparty"): .Orientation = xlRowField: .Position = 3: End With
        .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount", xlSum
        .PivotFields("Sum of Transaction Amount").NumberFormat = "$#,#00.00"
        .AddDataField .PivotFields("Transaction Amount"), "Count of Transaction Amount", xlCount
        .RowAxisLayout xlCompactRow
        .PivotFields("Count of Transaction Amount").NumberFormat = "0"
        .PivotFields("Alert Information").AutoSort xlDescending, "Sum of Transaction Amount"
        .PivotFields("Counterparty").AutoSort xlDescending, "Sum of Transaction Amount"
    End With
    On Error GoTo 0

    ' The date-grouped pivots can't group a Transaction Date column that has
    ' blanks in it, so blank-date rows are dropped from a hidden copy of the
    ' data first. With no blanks - always the case for Lookback, where every
    ' row was kept for its date - that copy would be identical to the data
    ' sheet, so it is skipped and these pivots read the data sheet directly.
    ' At 600,000 rows the copy is another full set of every row in memory.
    Dim needScratch As Boolean, wsGrp As Worksheet
    needScratch = False
    If dCol > 0 Then
        ddLast = wsData.Cells(wsData.Rows.count, dCol).End(xlUp).row
        If ddLast > 1 Then
            needScratch = (Application.WorksheetFunction.CountBlank( _
                wsData.Range(wsData.Cells(2, dCol), wsData.Cells(ddLast, dCol))) > 0)
        End If
    End If

    If needScratch Then
        On Error Resume Next
        Set wsPvScratch = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
        wsPvScratch.Visible = xlSheetVeryHidden
        If wsData.UsedRange.Cells.count > 0 Then
            wsData.UsedRange.Copy Destination:=wsPvScratch.Range("A1")
        End If
        On Error GoTo 0
        If wsPvScratch Is Nothing Then GoTo SkipDateGroupedPivots

        On Error Resume Next
        ddLast = wsPvScratch.Cells(wsPvScratch.Rows.count, dCol).End(xlUp).row
        If ddLast > 1 Then
            wsPvScratch.Range(wsPvScratch.Cells(2, dCol), wsPvScratch.Cells(ddLast, dCol)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            wsPvScratch.Range(wsPvScratch.Cells(2, dCol), wsPvScratch.Cells(ddLast, dCol)).NumberFormat = "m/d/yyyy"
        End If
        On Error GoTo 0
        Set wsGrp = wsPvScratch
    Else
        Set wsGrp = wsData
    End If

    lastRow = wsGrp.Cells(wsGrp.Rows.count, "A").End(xlUp).row
    lastCol = wsGrp.Cells(1, wsGrp.Columns.count).End(xlToLeft).Column
    If lastRow > 1 Then
        Set rngTmp = wsGrp.Range(wsGrp.Cells(1, 1), wsGrp.Cells(lastRow, lastCol))
        Set cacheTmp = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PivotSourceAddress(rngTmp))

        ' PIVOT 2: TEMPORAL
        Set ptx = cacheTmp.CreatePivotTable(TableDestination:=wsPv.Range("F3"), TableName:="TemporalPivot")
        On Error Resume Next
        With ptx
            .TableStyle2 = "PivotStyleLight16"
            With .PivotFields("Transaction Date"): .Orientation = xlRowField: .Position = 1: End With
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount ", xlSum
            .PivotFields("Sum of Transaction Amount ").NumberFormat = "$#,#00.00"
            .AddDataField .PivotFields("Transaction Amount"), "Count of Transaction Amount ", xlCount
            .PivotFields("Count of Transaction Amount ").NumberFormat = "0"
        End With
        wsPv.Range("F4").Group Start:=True, End:=True, Periods:=Array(False, False, False, True, True, False, True)
        On Error Resume Next
        ptx.PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        On Error GoTo 0

        ' PIVOT 3: ENHANCED TEMPORAL
        ptx.TableRange2.Copy Destination:=wsPv.Range("K3")
        Set ptE = wsPv.Range("K3").PivotTable
        ptE.Name = "TemporalPivot_Enhanced"
        On Error Resume Next
        With ptE
            .PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
            .PivotFields("Count of Transaction Amount ").Orientation = xlHidden
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 4: End With
            .AddDataField .PivotFields("Transaction Amount"), "No of Trx  ", xlCount
            .PivotFields("No of Trx  ").NumberFormat = "0"
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount  ", xlSum
            .PivotFields("Sum of Transaction Amount  ").NumberFormat = "$#,#00.00"
            .PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        End With
        On Error GoTo 0

        ' PIVOT 4: DR/CR INVERTED
        ptx.TableRange2.Copy Destination:=wsPv.Range("Q3")
        Set ptD = wsPv.Range("Q3").PivotTable
        ptD.Name = "DrCrTemporalPivot"
        On Error Resume Next
        With ptD
            .PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
            .PivotFields("Count of Transaction Amount ").Orientation = xlHidden
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 1: End With
            .AddDataField .PivotFields("Transaction Amount"), "No of Trx   ", xlCount
            .PivotFields("No of Trx   ").NumberFormat = "0"
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount   ", xlSum
            .PivotFields("Sum of Transaction Amount   ").NumberFormat = "$#,#00.00"
            .PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        End With
        On Error GoTo 0

        On Error Resume Next
        BulletproofDateFormat wsPv
        HighlightDrCrRows ptE, wsPv
        HighlightDrCrRows ptD, wsPv
        On Error GoTo 0
    End If

SkipDateGroupedPivots:
    SafeDeleteSheet wb, "TempPivotScratch"
    wsPv.Columns("A:W").AutoFit
End Sub

' ==========================================================
' SetSheetZoom85 - forces each named sheet's saved window zoom to 85%,
' matching ConsolidatedData's OWN saved zoom in the live workbook (its
' sheetView zoomScale is 85, not 100 - confirmed directly in the file's
' XML). A sheet built via Sheets.Add inside a brand-new Workbooks.Add
' workbook defaults to 100% instead of inheriting whatever zoom the
' analyst's own window happened to be at - which is what the OLD
' Worksheet.Copy-based approach used to carry over silently, and is the
' most likely reason exported sheets used to look "different" even
' though the font itself was never actually changed by that rewrite.
'
' Zoom is a WINDOW property in VBA (Window.Zoom), not a range/cell
' property, so each sheet must be briefly activated to set it. This
' MUST run BEFORE SaveAs - Zoom set after a workbook is already written
' to disk never makes it into that saved file.
' ==========================================================
Private Sub SetSheetZoom85(ByVal wb As Workbook, ByVal sheetNames As Variant)
    On Error Resume Next
    Dim nm As Variant
    For Each nm In sheetNames
        wb.Sheets(CStr(nm)).Activate
        ActiveWindow.Zoom = 85
    Next nm
    On Error GoTo 0
End Sub

' ==========================================================
' MoveExistingExportAside - never overwrite an export in place.
' ==========================================================
' Evidence from a real run: the Lookback file saved and the Combined file
' did not, from the same code, into the same folder. The difference was
' that the Lookback had a NEW filename (its date range had just changed),
' so SaveAs created a fresh file - while the Combined filename is identical
' on every run, so SaveAs had to OVERWRITE the copy left by the last run.
'
' Excel overwrites by writing a temporary file and then swapping it in for
' the old one. On the OneDrive-synced Desktop that replace is exactly what
' sync interferes with, and it fails as a bare "Method 'SaveAs' of object
' '_Workbook' failed". Creating a new file does not hit it - which is why
' the Lookback succeeded.
'
' So the old export is renamed out of the way first, and every save is a
' plain create. The old copy is kept as "<name> (previous).xlsx" until the
' new save has definitely succeeded, then deleted - so if anything still
' goes wrong, the previous export is still on disk under a readable name
' rather than lost.
'
' (This replaces an earlier check that opened the existing file for
' exclusive write access to test for a lock. That was the wrong test: the
' open could pass while the replace still failed, and on a synced folder
' opening the file for write can itself prompt OneDrive to pick it up.)
Private Sub MoveExistingExportAside(ByVal targetPath As String)
    Dim asidePath As String, moved As Boolean
    asidePath = PreviousExportPath(targetPath)

    On Error Resume Next
    If Len(Dir(targetPath)) = 0 Then Exit Sub          ' nothing to move
    If Len(Dir(asidePath)) > 0 Then Kill asidePath       ' stale copy from a failed run
    Err.Clear
    Name targetPath As asidePath
    moved = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If Not moved Then
        Err.Raise vbObjectError + 1004, "Consolidated_AML_Workflow", _
            "The export could not replace the existing file, because it is open " & _
            "or locked:" & vbCrLf & vbCrLf & _
            Mid$(targetPath, InStrRev(targetPath, Application.PathSeparator) + 1) & _
            vbCrLf & vbCrLf & _
            "Close every open copy of it - check all Excel windows, and Task " & _
            "Manager for a second EXCEL.EXE - or wait for OneDrive to finish " & _
            "syncing it. Then run the export again."
    End If
End Sub

' Called only AFTER a SaveAs has succeeded, so the previous copy is never
' removed until its replacement exists.
Private Sub DiscardPreviousExport(ByVal targetPath As String)
    On Error Resume Next
    Dim asidePath As String
    asidePath = PreviousExportPath(targetPath)
    If Len(Dir(asidePath)) > 0 Then Kill asidePath
    On Error GoTo 0
End Sub

' "...\X.xlsx" -> "...\X (previous).xlsx"
Private Function PreviousExportPath(ByVal targetPath As String) As String
    Dim dot As Long
    dot = InStrRev(targetPath, ".")
    If dot = 0 Then
        PreviousExportPath = targetPath & " (previous)"
    Else
        PreviousExportPath = Left$(targetPath, dot - 1) & " (previous)" & Mid$(targetPath, dot)
    End If
End Function

' ==========================================================
' CloseIfAlreadyOpen - safely closes an open target file prior to SaveAs
' ==========================================================
Private Sub CloseIfAlreadyOpen(ByVal targetPath As String)
    On Error Resume Next
    Dim targetName As String, wb As Workbook
    targetName = Mid$(targetPath, InStrRev(targetPath, Application.PathSeparator) + 1)
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            If StrComp(wb.Name, targetName, vbTextCompare) = 0 Then
                Application.DisplayAlerts = False
                wb.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit For
            End If
        End If
    Next wb
    On Error GoTo 0
End Sub

' ==========================================================
' LastDataRow - finds the true last data row across all columns
' ==========================================================
Private Function LastDataRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If c Is Nothing Then
        LastDataRow = 1
    Else
        LastDataRow = c.row
    End If
End Function

' ==========================================================
' TidyDataSheet - the column/row fit every exported data sheet gets
' ==========================================================
' Up to TIDY_FULL_FIT_ROWS rows: exactly as before - fit the columns, wrap
' text, fit every row to its wrapped text, align to the top.
'
' Above that, the fitting is what hangs Excel. Fitting a column measures
' every cell in it and fitting a wrapped row measures every cell across it,
' so at 600,000 rows x 38 columns each pass measures 22 million cells, on
' every sheet it runs on. Big sheets fit their columns to the first
' TIDY_SAMPLE_ROWS rows and stay unwrapped at standard row height.
' ==========================================================
Private Sub TidyDataSheet(ByVal ws As Worksheet)
    If LastDataRow(ws) <= TIDY_FULL_FIT_ROWS Then
        With ws.Cells
            .WrapText = False
            .EntireColumn.AutoFit
            .WrapText = True
            .EntireRow.AutoFit
            .VerticalAlignment = xlTop
        End With
    Else
        ws.Cells.WrapText = False
        ws.Range(ws.Rows(1), ws.Rows(TIDY_SAMPLE_ROWS)).Columns.AutoFit
        ws.Cells.VerticalAlignment = xlTop
    End If
End Sub

' ==========================================================
' Export workbooks - tracked until saved, closed unsaved if the run fails
' ==========================================================
Private Function NewOutputWorkbook() As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add
    m_unsaved.Add wb
    Set NewOutputWorkbook = wb
End Function

Private Sub MarkSaved(ByVal wb As Workbook)
    Dim i As Long
    For i = m_unsaved.count To 1 Step -1
        If m_unsaved(i) Is wb Then m_unsaved.Remove i
    Next i
End Sub

Private Sub CloseUnsavedExports()
    Dim wb As Variant
    On Error Resume Next
    If m_unsaved Is Nothing Then Exit Sub
    For Each wb In m_unsaved
        wb.Close SaveChanges:=False
    Next wb
    Set m_unsaved = New Collection
    On Error GoTo 0
End Sub

' ==========================================================
' FinishRun - hands Excel back after a successful export
' ==========================================================
Private Sub FinishRun(ByVal origCalc As XlCalculation)
    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = origCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    On Error Resume Next
    ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
    ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
    ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
    Application.OnTime Now + TimeSerial(0, 0, 1), "PushTrxTracker_Deferred"
    On Error GoTo 0
End Sub

' ==========================================================
' LoadQueryToSheet - runs a query into A1 of a sheet and leaves its rows
' there as a plain range: no table, no query link. Returns the row count.
' ==========================================================
Private Function LoadQueryToSheet(ByVal ws As Worksheet, ByVal queryName As String) As Long
    Dim lo As ListObject, n As Long

    ' SafeDeleteSheet turns alerts back on; a load must not stop on a dialog.
    Application.DisplayAlerts = False

    Set lo = ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & _
        ";Extended Properties=""""", Destination:=ws.Range("A1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .PreserveColumnInfo = False
        .Refresh BackgroundQuery:=False
    End With
    n = lo.ListRows.count

    ' A sheet holds 1,048,576 rows. A query with more than that is cut off
    ' at the bottom of the sheet with nothing more than a warning, which
    ' would drop transactions without anyone noticing - so a full sheet is
    ' treated as an error, never as a result.
    If n >= ws.Rows.count - 1 Then
        Err.Raise vbObjectError + 1004, "modLargeExport", _
            "The " & ws.Name & " sheet needs more rows than one Excel sheet holds (" & _
            Format$(ws.Rows.count - 1, "#,##0") & "), so it would have been cut short. " & _
            "Split the source files into smaller date ranges and export each range separately."
    End If

    ' Table Style Medium 2 is the banded blue look the source files carry into
    ' Module9's exports. Unlinking and converting to a range keeps that
    ' formatting but drops the query link, so the export holds plain data.
    lo.TableStyle = "TableStyleMedium2"
    On Error Resume Next
    lo.QueryTable.Delete
    If Err.Number <> 0 Then Err.Clear: lo.Unlink
    On Error GoTo 0
    lo.Unlist

    LoadQueryToSheet = n
End Function

' ==========================================================
' LoadQueryAcrossSheets - LoadQueryToSheet for a sheet that can outgrow
' Excel's row limit (Raw Transactions, CP Selection, DeDupe).
' RAW_ROWS_PER_SHEET rows go on firstSheet, the next lot on "<name> (2)"
' right after it, and so on. Returns the total number of rows loaded.
' ==========================================================
Private Function LoadQueryAcrossSheets(ByVal wb As Workbook, ByVal firstSheet As Worksheet, _
    ByVal queryName As String) As Long
    Dim part As Long, n As Long, total As Long, ws As Worksheet, partQuery As String

    Set ws = firstSheet
    Do
        part = part + 1
        partQuery = queryName & "_Part" & part
        wb.Queries.Add Name:=partQuery, _
            Formula:=BuildRowRangeM(queryName, (part - 1) * RAW_ROWS_PER_SHEET, RAW_ROWS_PER_SHEET)
        n = LoadQueryToSheet(ws, partQuery)
        total = total + n

        ' Row count an exact multiple of the sheet size: the last sheet came
        ' back empty, so it goes.
        If n = 0 And part > 1 Then
            Application.DisplayAlerts = False
            ws.Delete
            Exit Do
        End If
        If n < RAW_ROWS_PER_SHEET Then Exit Do

        SetStage "loading " & firstSheet.Name & " (part " & (part + 1) & ")"
        Set ws = wb.Sheets.Add(After:=ws)
        ws.Name = firstSheet.Name & " (" & (part + 1) & ")"
    Loop

    LoadQueryAcrossSheets = total
End Function

' baseName itself, or one of its overflow sheets "baseName (2)", "(3)"...
Private Function IsSheetPart(ByVal nm As String, ByVal baseName As String) As Boolean
    IsSheetPart = (nm = baseName) Or (nm Like baseName & " (#)") Or (nm Like baseName & " (##)")
End Function

' Module9 sets 85% zoom on each named export sheet; with overflow sheets
' the names aren't fixed, so here it is every sheet.
Private Sub ZoomAllSheets(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        SetSheetZoom85 wb, Array(ws.Name)
    Next ws
End Sub

' ==========================================================
' RemoveAllQueries - strips every query and connection from an export
' workbook before it is saved (the data already loaded stays).
' ==========================================================
Private Sub RemoveAllQueries(ByVal wb As Workbook)
    Dim i As Long, pass As Long, sh As Worksheet
    On Error Resume Next
    ' Any query table LoadQueryToSheet couldn't unlink (the data stays).
    For Each sh In wb.Worksheets
        For i = sh.QueryTables.count To 1 Step -1
            sh.QueryTables(i).Delete
        Next i
    Next sh
    For i = wb.Connections.count To 1 Step -1
        wb.Connections(i).Delete
    Next i
    ' A query that another one refers to may refuse to go first, so a few
    ' passes, newest first.
    For pass = 1 To 5
        For i = wb.Queries.count To 1 Step -1
            wb.Queries(i).Delete
        Next i
        If wb.Queries.count = 0 Then Exit For
    Next pass
    On Error GoTo 0
End Sub

' ==========================================================
' ReadHeaderNames - the combined data's column names, in order
' ==========================================================
' Loads a one-column list of TrxBase's headers onto a scratch sheet, reads
' it and throws the sheet away. Returns Empty when no file had a
' "Transaction ID" header row.
' ==========================================================
Private Function ReadHeaderNames(ByVal wb As Workbook) As Variant
    Dim wsTmp As Worksheet, n As Long, v As Variant, names() As String, i As Long

    wb.Queries.Add Name:="TrxHeaders", Formula:=BuildHeadersM()
    Set wsTmp = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
    n = LoadQueryToSheet(wsTmp, "TrxHeaders")
    If n > 0 Then
        v = wsTmp.Range("A2").Resize(n, 1).Value
        ReDim names(1 To n)
        If n = 1 Then
            names(1) = CStr(v)
        Else
            For i = 1 To n
                names(i) = CStr(v(i, 1))
            Next i
        End If
        ReadHeaderNames = names
    End If

    Application.DisplayAlerts = False
    wsTmp.Delete
    On Error Resume Next
    wb.Queries("TrxHeaders").Delete
    On Error GoTo 0
End Function

' ==========================================================
' ResolveColumns / FindHeaderName - Module9's header lookups, on the names
' ==========================================================
' Module9 finds each column with Rows(1).Find: case-insensitive, "contains"
' except for the alerted flag, which must match whole. Find starts AFTER
' A1 and wraps round to it last, so FindHeaderName checks the names in
' that same order and picks the same column when more than one matches.
' ==========================================================
Private Sub ResolveColumns(ByRef hdr As Variant, ByRef colDate As String, ByRef colAmt As String, _
    ByRef colDrCr As String, ByRef colBen As String, ByRef colOrig As String, _
    ByRef colFlag As String, ByRef colTrans As String, ByRef colAlert As String, ByRef colCp As String)

    colDate = FindHeaderName(hdr, "Transaction Date", False)
    colAmt = FindHeaderName(hdr, "Transaction Amount", False)
    colDrCr = FindHeaderName(hdr, "Dr Cr", False)
    colBen = FindHeaderName(hdr, "Beneficiary Name", False)
    colOrig = FindHeaderName(hdr, "Originator Name", False)
    colFlag = FindHeaderName(hdr, "Is Alerted Transaction?", True)
    colTrans = FindHeaderName(hdr, "Transaction ID", False)
    colAlert = FindHeaderName(hdr, "Alert Information", False)

    ' Power Query can't add a column under a name that's already taken.
    colCp = "Counterparty"
    If Len(FindHeaderName(hdr, colCp, True)) > 0 Then colCp = "Counterparty (calculated)"
End Sub

Private Function FindHeaderName(ByRef names As Variant, ByVal what As String, ByVal wholeMatch As Boolean) As String
    Dim n As Long, k As Long, i As Long, nm As String
    n = UBound(names) - LBound(names) + 1
    For k = 1 To n
        i = LBound(names) + (k Mod n)     ' 2nd, 3rd, ... last, then 1st
        nm = CStr(names(i))
        If wholeMatch Then
            If StrComp(nm, what, vbTextCompare) = 0 Then FindHeaderName = nm: Exit Function
        Else
            If InStr(1, nm, what, vbTextCompare) > 0 Then FindHeaderName = nm: Exit Function
        End If
    Next k
End Function

Private Function NoHeaderMessage(ByVal folderPath As String) As String
    NoHeaderMessage = "None of the Excel files in:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
        "has a 'Transaction ID' header in its first 100 rows, so there is nothing to export."
End Function

' ==========================================================
' Sheet helpers for the loaded data
' ==========================================================
' Column number of a header in row 1 (exact name, any case), or 0.
Private Function HeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastC As Long, c As Long
    If Len(headerName) = 0 Then Exit Function
    lastC = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastC
        If StrComp(CStr(ws.Cells(1, c).Value), headerName, vbTextCompare) = 0 Then
            HeaderColumn = c
            Exit Function
        End If
    Next c
End Function

' Sets the date and amount number formats on a loaded sheet. An empty
' format or name skips that column.
Private Sub FormatDataColumns(ByVal ws As Worksheet, ByVal dateName As String, _
    ByVal dateFormat As String, ByVal amtName As String)
    Dim lastR As Long, c As Long
    lastR = LastDataRow(ws)
    If lastR < 2 Then Exit Sub
    If Len(dateFormat) > 0 Then
        c = HeaderColumn(ws, dateName)
        If c > 0 Then ws.Range(ws.Cells(2, c), ws.Cells(lastR, c)).NumberFormat = dateFormat
    End If
    c = HeaderColumn(ws, amtName)
    If c > 0 Then ws.Range(ws.Cells(2, c), ws.Cells(lastR, c)).NumberFormat = "$#,##0.00"
End Sub

' Earliest and latest date on the loaded Alerted sheet. False when it has
' no dated rows (or no date column) - Module9's "No dated 'Yes' alerted
' transactions" case.
Private Function AlertedDateSpan(ByVal ws As Worksheet, ByVal dateName As String, _
    ByRef firstD As Date, ByRef lastD As Date) As Boolean
    Dim dc As Long, lastR As Long, v As Variant, r As Long, have As Boolean, d As Date
    dc = HeaderColumn(ws, dateName)
    If dc = 0 Then Exit Function
    lastR = LastDataRow(ws)
    If lastR < 2 Then Exit Function

    v = ws.Range(ws.Cells(2, dc), ws.Cells(lastR, dc)).Value
    If Not IsArray(v) Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = v
        v = one
    End If

    For r = 1 To UBound(v, 1)
        If VarType(v(r, 1)) = vbDate Then
            d = v(r, 1)
            If Not have Then
                firstD = d: lastD = d: have = True
            Else
                If d < firstD Then firstD = d
                If d > lastD Then lastD = d
            End If
        End If
    Next r
    AlertedDateSpan = have
End Function

' ==========================================================
' Power Query (M) text
' ==========================================================
' M is written here with ` in place of " so the VBA stays readable, and
' MTpl turns them back. Anything that comes from the data or the file
' system - folder, file and column names - goes in through MText, which
' quotes it properly, never through MTpl.
' ==========================================================
Private Sub AddLine(ByRef m As String, ByVal s As String)
    m = m & s & vbLf
End Sub

Private Function MTpl(ByVal s As String) As String
    MTpl = Replace(s, "`", """")
End Function

Private Function MText(ByVal s As String) As String
    MText = """" & Replace(s, """", """""") & """"
End Function

Private Function MDate(ByVal d As Date) As String
    MDate = "#date(" & Year(d) & ", " & Month(d) & ", " & Day(d) & ")"
End Function

' TrxBase - every Excel file directly in the folder, combined.
' Each file: its first sheet, from the row and column of its "Transaction
' ID" header (whole cell, any case - Module9's Find) across and down. Files
' without that header are skipped, as in Module9; a file that can't be read
' at all stops the export with its name. Files are taken in name order and
' matched up by column name. Rows with nothing in them are dropped.
Private Function BuildBaseM(ByVal folderWithSlash As String, ByVal excluded As Collection) As String
    Dim m As String, ex As String, item As Variant
    For Each item In excluded
        If Len(ex) > 0 Then ex = ex & ", "
        ex = ex & MText(LCase$(CStr(item)))
    Next item

    AddLine m, "let"
    AddLine m, "    FolderPath = " & MText(folderWithSlash) & ","
    AddLine m, "    Excluded = {" & ex & "},"
    AddLine m, MTpl("    IsHeaderCell = (v as any) as logical => v is text and Text.Lower(v) = `transaction id`,")
    AddLine m, "    HeaderPos = (raw as table) as nullable record =>"
    AddLine m, "        let"
    AddLine m, "            Top = Table.ToRows(Table.FirstN(raw, 100)),"
    AddLine m, "            Hits = List.Select(List.Positions(Top), (i) => List.AnyTrue(List.Transform(Top{i}, IsHeaderCell))),"
    AddLine m, "            R = if List.IsEmpty(Hits) then null else Hits{0}"
    AddLine m, "        in"
    AddLine m, "            if R = null then null else [Row = R, Col = List.PositionOf(List.Transform(Top{R}, IsHeaderCell), true)],"
    AddLine m, "    LoadFile = (content as binary) as nullable table =>"
    AddLine m, "        let"
    AddLine m, "            Book = Excel.Workbook(content, null, true),"
    AddLine m, MTpl("            Sheets = Table.SelectRows(Book, each [Kind] = `Sheet`),")
    AddLine m, "            Raw = if Table.IsEmpty(Sheets) then null else Sheets{0}[Data],"
    AddLine m, "            Pos = if Raw = null then null else HeaderPos(Raw),"
    AddLine m, "            Body = if Pos = null then null else Table.PromoteHeaders(Table.Skip(Raw, Pos[Row]), [PromoteAllScalars = true])"
    AddLine m, "        in"
    AddLine m, "            if Body = null then null else Table.RemoveColumns(Body, List.FirstN(Table.ColumnNames(Body), Pos[Col])),"
    AddLine m, "    ReadFile = (content as binary, name as text) as nullable table =>"
    AddLine m, MTpl("        try LoadFile(content) otherwise error Error.Record(`DataFormat.Error`, `Could not read ` & name & ` as an Excel workbook.`),")
    AddLine m, MTpl("    IsBlankRow = (r as record) as logical => List.AllTrue(List.Transform(Record.FieldValues(r), each _ = null or _ = ``)),")
    AddLine m, "    Files = Folder.Files(FolderPath),"
    AddLine m, MTpl("    Picked = Table.SelectRows(Files, each Text.Lower(Text.TrimEnd([Folder Path], `\`)) = Text.Lower(Text.TrimEnd(FolderPath, `\`))")
    AddLine m, MTpl("        and Text.Contains(Text.Lower([Name]), `.xls`)")
    AddLine m, MTpl("        and not Text.StartsWith([Name], `~$`)")
    AddLine m, "        and not List.Contains(Excluded, Text.Lower([Name]))),"
    AddLine m, MTpl("    Keyed = Table.AddColumn(Picked, `SortKey`, each Text.Upper([Name])),")
    AddLine m, MTpl("    Ordered = Table.Sort(Keyed, {{`SortKey`, Order.Ascending}}),")
    AddLine m, "    Tables = List.Transform(Table.ToRecords(Ordered), each ReadFile([Content], [Name])),"
    AddLine m, "    Combined = Table.Combine(List.RemoveNulls(Tables)),"
    AddLine m, "    Result = Table.SelectRows(Combined, each not IsBlankRow(_))"
    AddLine m, "in"
    AddLine m, "    Result"
    BuildBaseM = m
End Function

' Rows offset+1 .. offset+count of sourceQuery (fewer, or none, at the end).
Private Function BuildRowRangeM(ByVal sourceQuery As String, ByVal offset As Long, ByVal count As Long) As String
    Dim m As String
    AddLine m, "let"
    AddLine m, "    Source = " & sourceQuery & ","
    AddLine m, "    Result = Table.FirstN(Table.Skip(Source, " & offset & "), " & count & ")"
    AddLine m, "in"
    AddLine m, "    Result"
    BuildRowRangeM = m
End Function

' TrxHeaders - TrxBase's column names as a one-column table.
Private Function BuildHeadersM() As String
    Dim m As String
    AddLine m, "let"
    AddLine m, "    Source = TrxBase,"
    AddLine m, MTpl("    Result = Table.FromColumns({Table.ColumnNames(Source)}, {`Header`})")
    AddLine m, "in"
    AddLine m, "    Result"
    BuildHeadersM = m
End Function

' TrxClean - Module9's step 3 cleanup on TrxBase:
'   - Transaction Date: text dates read as month/day/year (Module9's
'     TextToColumns MDY); midnight date-times become plain dates.
'     Anything that isn't a date is left as it was.
'   - Transaction Amount: numbers held as text become numbers.
'   - Counterparty: Beneficiary Name on DR rows, Originator Name
'     otherwise, blank when that name is blank - Module9's formula, as a
'     value. Only added when all three columns exist, as in Module9.
Private Function BuildCleanM(ByVal colDate As String, ByVal colAmt As String, ByVal colDrCr As String, _
    ByVal colBen As String, ByVal colOrig As String, ByVal colCp As String) As String
    Dim m As String, prevStep As String

    AddLine m, "let"
    AddLine m, "    Source = TrxBase,"
    prevStep = "Source"

    If Len(colDate) > 0 Then
        AddLine m, "    AsDate = (v as any) as any =>"
        AddLine m, MTpl("        if v = null or v = `` then null")
        AddLine m, "        else if v is date then v"
        AddLine m, "        else if v is datetime then (if DateTime.Time(v) = #time(0, 0, 0) then DateTime.Date(v) else v)"
        AddLine m, "        else if v is number then (try (if v = Number.RoundDown(v) then Date.From(v) else DateTime.From(v)) otherwise v)"
        AddLine m, "        else if v is text then ("
        AddLine m, "            let"
        AddLine m, "                t = Text.Trim(v),"
        AddLine m, MTpl("                p = if t = `` then null else (try DateTime.From(t, `en-US`) otherwise null)")
        AddLine m, "            in"
        AddLine m, "                if p = null then v"
        AddLine m, "                else if DateTime.Time(p) = #time(0, 0, 0) then DateTime.Date(p)"
        AddLine m, "                else p)"
        AddLine m, "        else v,"
        AddLine m, "    Dated = Table.TransformColumns(" & prevStep & ", {{" & MText(colDate) & ", AsDate}}),"
        prevStep = "Dated"
    End If

    If Len(colAmt) > 0 Then
        AddLine m, "    AsAmount = (v as any) as any =>"
        AddLine m, MTpl("        if v = `` then null")
        AddLine m, "        else if v is text then ("
        AddLine m, "            let"
        AddLine m, "                t = Text.Trim(v),"
        AddLine m, MTpl("                p = if t = `` then null else (try Number.From(t, `en-US`) otherwise (try Number.From(Text.Remove(t, {`$`, `,`}), `en-US`) otherwise null))")
        AddLine m, "            in"
        AddLine m, "                if p = null then v else p)"
        AddLine m, "        else v,"
        AddLine m, "    Amounts = Table.TransformColumns(" & prevStep & ", {{" & MText(colAmt) & ", AsAmount}}),"
        prevStep = "Amounts"
    End If

    If Len(colDrCr) > 0 And Len(colBen) > 0 And Len(colOrig) > 0 Then
        AddLine m, "    WithCounterparty = Table.AddColumn(" & prevStep & ", " & MText(colCp) & ", each"
        AddLine m, "        let"
        AddLine m, "            dr = Record.Field(_, " & MText(colDrCr) & "),"
        AddLine m, "            ben = Record.Field(_, " & MText(colBen) & "),"
        AddLine m, "            org = Record.Field(_, " & MText(colOrig) & ")"
        AddLine m, "        in"
        AddLine m, MTpl("            if dr is text and Text.Upper(dr) = `DR` then (if ben = null or ben = `` then `` else ben)")
        AddLine m, MTpl("            else (if org = null or org = `` then `` else org)),")
        prevStep = "WithCounterparty"
    End If

    AddLine m, "    Result = " & prevStep
    AddLine m, "in"
    AddLine m, "    Result"
    BuildCleanM = m
End Function

' Rows of sourceQuery whose flag column (spaces trimmed, exact case, as
' Module9's Trim(CStr(...)) = "Yes") equals flagWant, and/or whose date
' falls inside winStart..winEnd. A date-time counts by its date, so a
' transaction at 3 pm on the last day of the window is inside it.
Private Function BuildFilterM(ByVal sourceQuery As String, ByVal flagCol As String, ByVal flagWant As String, _
    ByVal dateCol As String, ByVal useWindow As Boolean, ByVal winStart As Date, ByVal winEnd As Date) As String
    Dim m As String, cond As String

    AddLine m, "let"
    AddLine m, "    Source = " & sourceQuery & ","
    AddLine m, MTpl("    IsFlag = (v as any, want as text) as logical => v is text and Text.Trim(v, ` `) = want,")
    AddLine m, "    AsDay = (v as any) as nullable date => if v is date then v else if v is datetime then DateTime.Date(v) else null,"
    AddLine m, "    WinStart = " & MDate(winStart) & ","
    AddLine m, "    WinEnd = " & MDate(winEnd) & ","
    AddLine m, "    InWindow = (v as any) as logical => (let d = AsDay(v) in d <> null and d >= WinStart and d <= WinEnd),"

    If Len(flagCol) > 0 Then cond = "IsFlag(Record.Field(_, " & MText(flagCol) & "), " & MText(flagWant) & ")"
    If useWindow Then
        If Len(cond) > 0 Then cond = cond & " and "
        cond = cond & "InWindow(Record.Field(_, " & MText(dateCol) & "))"
    End If

    AddLine m, "    Result = Table.SelectRows(Source, each " & cond & ")"
    AddLine m, "in"
    AddLine m, "    Result"
    BuildFilterM = m
End Function

' sourceQuery less duplicate rows on keyCols (the first one kept, as
' RemoveDuplicates does), then less rows with no value in blankDateCol.
' keyCols = Empty and/or blankDateCol = "" skip that step.
'
' Duplicates are matched on each key as upper-case text, so 1877576904 and
' "1877576904" - the same ID from two files, one holding it as a number
' and one as text - are one transaction, and case doesn't matter. That is
' how Module9's RemoveDuplicates saw them, after its value rewrite had
' turned both into the same number.
Private Function BuildDedupeM(ByVal sourceQuery As String, ByVal keyCols As Variant, _
    ByVal blankDateCol As String) As String
    Dim m As String, prevStep As String, keys As String, k As Variant, i As Long

    AddLine m, "let"
    AddLine m, "    Source = " & sourceQuery & ","
    AddLine m, MTpl("    IsBlankValue = (v as any) as logical => v = null or v = ``,")
    AddLine m, "    KeyText = (v as any) as nullable text => if v = null then null else Text.Upper(Text.From(v)),"
    prevStep = "Source"

    If IsArray(keyCols) Then
        For Each k In keyCols
            i = i + 1
            AddLine m, "    Keyed" & i & " = Table.AddColumn(" & prevStep & ", " & MText("__DupKey" & i) & _
                ", each KeyText(Record.Field(_, " & MText(CStr(k)) & "))),"
            prevStep = "Keyed" & i
            If Len(keys) > 0 Then keys = keys & ", "
            keys = keys & MText("__DupKey" & i)
        Next k
        AddLine m, "    Distinct = Table.RemoveColumns(Table.Distinct(" & prevStep & ", {" & keys & "}), {" & keys & "}),"
        prevStep = "Distinct"
    End If

    If Len(blankDateCol) > 0 Then
        AddLine m, "    Dated = Table.SelectRows(" & prevStep & ", each not IsBlankValue(Record.Field(_, " & MText(blankDateCol) & "))),"
        prevStep = "Dated"
    End If

    AddLine m, "    Result = " & prevStep
    AddLine m, "in"
    AddLine m, "    Result"
    BuildDedupeM = m
End Function

' ==========================================================
' ReadQueryRow - the first row of a query's result, as an array (1 To
' columns), via a scratch sheet. Empty if the query returned no rows.
' ==========================================================
Private Function ReadQueryRow(ByVal wb As Workbook, ByVal queryName As String) As Variant
    Dim wsTmp As Worksheet, lastC As Long, out() As Variant, c As Long
    Set wsTmp = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
    If LoadQueryToSheet(wsTmp, queryName) > 0 Then
        lastC = wsTmp.Cells(1, wsTmp.Columns.count).End(xlToLeft).Column
        ReDim out(1 To lastC)
        For c = 1 To lastC
            out(c) = wsTmp.Cells(2, c).Value
        Next c
        ReadQueryRow = out
    End If
    Application.DisplayAlerts = False
    wsTmp.Delete
End Function

' ==========================================================
' BuildTotalsPivots - Legacy's four pivots, on the totals sheets
' ==========================================================
' Same positions, names, styles and number formats as Module9's Legacy
' pivots. Each source row is already a total (Transaction Amount = the sum,
' Transaction Count = how many), so every value field sums: the "Count of"
' fields sum Transaction Count, which is what counting the rows gave.
' ==========================================================
Private Sub BuildTotalsPivots(ByVal wb As Workbook, ByVal wsPv As Worksheet, _
    ByVal wsScn As Worksheet, ByVal wsDay As Worksheet)
    Dim lastR As Long, lastC As Long, cache As PivotCache
    Dim ptx As PivotTable, ptE As PivotTable, ptD As PivotTable

    ' PIVOT 1: SCENARIO
    lastR = LastDataRow(wsScn)
    lastC = wsScn.Cells(1, wsScn.Columns.count).End(xlToLeft).Column
    If lastR > 1 Then
        Set cache = wb.PivotCaches.Create(SourceType:=xlDatabase, _
            SourceData:=PivotSourceAddress(wsScn.Range(wsScn.Cells(1, 1), wsScn.Cells(lastR, lastC))))
        Set ptx = cache.CreatePivotTable(TableDestination:=wsPv.Range("A3"), TableName:="ScenarioPivot")
        On Error Resume Next
        With ptx
            .TableStyle2 = "PivotStyleLight16"
            With .PivotFields("Alert Information"): .Orientation = xlRowField: .Position = 1: End With
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 2: End With
            With .PivotFields("Counterparty"): .Orientation = xlRowField: .Position = 3: End With
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount", xlSum
            .PivotFields("Sum of Transaction Amount").NumberFormat = "$#,#00.00"
            .AddDataField .PivotFields("Transaction Count"), "Count of Transaction Amount", xlSum
            .RowAxisLayout xlCompactRow
            .PivotFields("Count of Transaction Amount").NumberFormat = "0"
            .PivotFields("Alert Information").AutoSort xlDescending, "Sum of Transaction Amount"
            .PivotFields("Counterparty").AutoSort xlDescending, "Sum of Transaction Amount"
        End With
        On Error GoTo 0
    End If

    ' PIVOTS 2-4: BY DATE
    lastR = LastDataRow(wsDay)
    lastC = wsDay.Cells(1, wsDay.Columns.count).End(xlToLeft).Column
    If lastR > 1 Then
        Set cache = wb.PivotCaches.Create(SourceType:=xlDatabase, _
            SourceData:=PivotSourceAddress(wsDay.Range(wsDay.Cells(1, 1), wsDay.Cells(lastR, lastC))))

        ' PIVOT 2: TEMPORAL
        Set ptx = cache.CreatePivotTable(TableDestination:=wsPv.Range("F3"), TableName:="TemporalPivot")
        On Error Resume Next
        With ptx
            .TableStyle2 = "PivotStyleLight16"
            With .PivotFields("Transaction Date"): .Orientation = xlRowField: .Position = 1: End With
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount ", xlSum
            .PivotFields("Sum of Transaction Amount ").NumberFormat = "$#,#00.00"
            .AddDataField .PivotFields("Transaction Count"), "Count of Transaction Amount ", xlSum
            .PivotFields("Count of Transaction Amount ").NumberFormat = "0"
        End With
        wsPv.Range("F4").Group Start:=True, End:=True, Periods:=Array(False, False, False, True, True, False, True)
        ptx.PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        On Error GoTo 0

        ' PIVOT 3: ENHANCED TEMPORAL
        ptx.TableRange2.Copy Destination:=wsPv.Range("K3")
        Set ptE = wsPv.Range("K3").PivotTable
        ptE.Name = "TemporalPivot_Enhanced"
        On Error Resume Next
        With ptE
            .PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
            .PivotFields("Count of Transaction Amount ").Orientation = xlHidden
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 4: End With
            .AddDataField .PivotFields("Transaction Count"), "No of Trx  ", xlSum
            .PivotFields("No of Trx  ").NumberFormat = "0"
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount  ", xlSum
            .PivotFields("Sum of Transaction Amount  ").NumberFormat = "$#,#00.00"
            .PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        End With
        On Error GoTo 0

        ' PIVOT 4: DR/CR INVERTED
        ptx.TableRange2.Copy Destination:=wsPv.Range("Q3")
        Set ptD = wsPv.Range("Q3").PivotTable
        ptD.Name = "DrCrTemporalPivot"
        On Error Resume Next
        With ptD
            .PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
            .PivotFields("Count of Transaction Amount ").Orientation = xlHidden
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 1: End With
            .AddDataField .PivotFields("Transaction Count"), "No of Trx   ", xlSum
            .PivotFields("No of Trx   ").NumberFormat = "0"
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount   ", xlSum
            .PivotFields("Sum of Transaction Amount   ").NumberFormat = "$#,#00.00"
            .PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        End With

        BulletproofDateFormat wsPv
        HighlightDrCrRows ptE, wsPv
        HighlightDrCrRows ptD, wsPv
        On Error GoTo 0
    End If

    wsPv.Columns("A:W").AutoFit
End Sub

' ==========================================================
' Sheet7's narrative figures for a case too big for ConsolidatedData
' ==========================================================
' Sheet7 works its figures out from ConsolidatedData with whole-column
' formulas (COUNT, SUM, MIN/MAX date, account list, MINIFS/MAXIFS,
' CR/DR SUMPRODUCT and COUNTIF). A case too big for ConsolidatedData gets
' those figures from Power Query instead (BuildStatsM), stored on the very
' hidden _LargeCaseStats sheet with that case's ECM ID in B1.
'
' Each such Sheet7 formula is wrapped, once, as
'     =IF(AND(<B1> <> "", <B1> = Sheet1!J9), <stored figure>, <original>)
' so the stored figure is used only while Sheet1 holds that ECM ID. Any
' other case - including the next one, or this one once Reset clears J9 -
' gets the original formula. A normal-size export clears B1 as well.
' ==========================================================
Private Function StatsSheet(ByVal createIfMissing As Boolean) As Worksheet
    On Error Resume Next
    Set StatsSheet = ThisWorkbook.Worksheets(STATS_SHEET)
    On Error GoTo 0
    If StatsSheet Is Nothing And createIfMissing Then
        Set StatsSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        StatsSheet.Name = STATS_SHEET
        StatsSheet.Visible = xlSheetVeryHidden
    End If
End Function

Private Sub ClearLargeCaseStats()
    Dim sh As Worksheet
    Set sh = StatsSheet(False)
    If Not sh Is Nothing Then sh.Range("B1").ClearContents
End Sub

' stats = ReadQueryRow of BuildStatsM: count, sum, first date, last date,
' accounts, smallest and largest amount over 0, CR sum, DR sum, CR count,
' DR count. Stored exactly as each Sheet7 formula shows it (its TEXT()
' format where it has one), with a leading ' so Excel keeps text as text.
Private Sub WriteLargeCaseStats(ByVal stats As Variant, ByVal ecmValue As Variant)
    Dim sh As Worksheet
    Set sh = StatsSheet(True)
    sh.Cells.Clear

    sh.Range("A1").Value = "ECM ID"
    ' Same type as Sheet1!J9, or the = in Sheet7 never matches.
    If VarType(ecmValue) = vbString Then
        sh.Range("B1").Value = "'" & ecmValue
    Else
        sh.Range("B1").Value = ecmValue
    End If

    sh.Range("A2").Value = "Transaction count":        sh.Range("B2").Value = NumOr0(stats(1))
    sh.Range("A3").Value = "Total amount":             sh.Range("B3").Value = "'" & Format$(NumOr0(stats(2)), "#,##0.00")
    sh.Range("A4").Value = "First date":               sh.Range("B4").Value = "'" & DateTextOf(stats(3))
    sh.Range("A5").Value = "Last date":                sh.Range("B5").Value = "'" & DateTextOf(stats(4))
    sh.Range("A6").Value = "Account numbers":          sh.Range("B6").Value = "'" & CStr(stats(5))
    sh.Range("A7").Value = "Smallest amount over 0":   sh.Range("B7").Value = "'" & Format$(NumOr0(stats(6)), "$#,##0.00")
    sh.Range("A8").Value = "Largest amount over 0":    sh.Range("B8").Value = "'" & Format$(NumOr0(stats(7)), "$#,##0.00")
    sh.Range("A9").Value = "CR total":                 sh.Range("B9").Value = "'" & Format$(NumOr0(stats(8)), "#,##0.00")
    sh.Range("A10").Value = "DR total":                sh.Range("B10").Value = "'" & Format$(NumOr0(stats(9)), "#,##0.00")
    sh.Range("A11").Value = "CR count":                sh.Range("B11").Value = NumOr0(stats(10))
    sh.Range("A12").Value = "DR count":                sh.Range("B12").Value = NumOr0(stats(11))
End Sub

Private Function NumOr0(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumOr0 = CDbl(v)
End Function

Private Function DateTextOf(ByVal v As Variant) As String
    If VarType(v) = vbDate Then
        DateTextOf = Format$(v, "mm/dd/yyyy")
    Else
        DateTextOf = CStr(v)
    End If
End Function

' Wraps each Sheet7 formula that reads ConsolidatedData (see above).
' Already-wrapped formulas are left alone. Returns the addresses of any it
' didn't recognise, or couldn't change, for the completion message.
Private Function PointSheet7AtLargeStats() As String
    Dim ws7 As Worksheet, c As Range, f As String, statRow As Long, unmatched As String, cond As String

    Set ws7 = ThisWorkbook.Worksheets("Sheet7")
    On Error Resume Next
    ws7.Unprotect Password:="p7ss"
    On Error GoTo 0

    cond = "AND('" & STATS_SHEET & "'!$B$1<>"""",'" & STATS_SHEET & "'!$B$1=Sheet1!$J$9)"
    For Each c In ws7.UsedRange
        If c.HasFormula Then
            f = c.Formula2
            If InStr(1, f, "ConsolidatedData", vbTextCompare) > 0 And _
               InStr(1, f, STATS_SHEET, vbTextCompare) = 0 Then
                statRow = StatRowForFormula(UCase$(f))
                If statRow > 0 Then
                    On Error Resume Next
                    c.Formula2 = "=IF(" & cond & ",'" & STATS_SHEET & "'!$B$" & statRow & "," & Mid$(f, 2) & ")"
                    If Err.Number <> 0 Then statRow = 0
                    On Error GoTo 0
                End If
                If statRow = 0 Then
                    If Len(unmatched) > 0 Then unmatched = unmatched & ", "
                    unmatched = unmatched & c.Address(False, False)
                End If
            End If
        End If
    Next c
    PointSheet7AtLargeStats = unmatched
End Function

' Which _LargeCaseStats row stands in for a Sheet7 formula (upper-cased).
' COUNTIF / SUMPRODUCT / MINIFS / MAXIFS are checked before COUNT( / SUM(
' / MIN( / MAX( so each formula lands on the right one.
Private Function StatRowForFormula(ByVal u As String) As Long
    If InStr(u, "TEXTJOIN") > 0 Then
        StatRowForFormula = 6
    ElseIf InStr(u, "MINIFS") > 0 Then
        StatRowForFormula = 7
    ElseIf InStr(u, "MAXIFS") > 0 Then
        StatRowForFormula = 8
    ElseIf InStr(u, "SUMPRODUCT") > 0 And InStr(u, """CR""") > 0 Then
        StatRowForFormula = 9
    ElseIf InStr(u, "SUMPRODUCT") > 0 And InStr(u, """DR""") > 0 Then
        StatRowForFormula = 10
    ElseIf InStr(u, "COUNTIF") > 0 And InStr(u, """CR""") > 0 Then
        StatRowForFormula = 11
    ElseIf InStr(u, "COUNTIF") > 0 And InStr(u, """DR""") > 0 Then
        StatRowForFormula = 12
    ElseIf InStr(u, "COUNT(") > 0 And InStr(u, "K:K") > 0 Then
        StatRowForFormula = 2
    ElseIf InStr(u, "SUM(") > 0 And InStr(u, "K:K") > 0 Then
        StatRowForFormula = 3
    ElseIf InStr(u, "MIN(") > 0 And InStr(u, "L:L") > 0 Then
        StatRowForFormula = 4
    ElseIf InStr(u, "MAX(") > 0 And InStr(u, "L:L") > 0 Then
        StatRowForFormula = 5
    End If
End Function

' Totals of sourceQuery for each combination of keyCols (named outNames in
' the result): Transaction Amount = sum of the numeric amounts, Transaction
' Count = rows with an amount (what a pivot's Count of Transaction Amount
' counts). dayCol, if given, is cut to its date first so totals are per
' day. Key columns that don't exist are left out.
Private Function BuildTotalsM(ByVal sourceQuery As String, ByVal keyCols As Variant, ByVal outNames As Variant, _
    ByVal amtCol As String, ByVal dayCol As String) As String
    Dim m As String, prevStep As String, i As Long
    Dim picks As String, renames As String, groupKeys As String

    For i = LBound(keyCols) To UBound(keyCols)
        If Len(keyCols(i)) > 0 Then
            If Len(picks) > 0 Then picks = picks & ", ": groupKeys = groupKeys & ", "
            picks = picks & MText(CStr(keyCols(i)))
            groupKeys = groupKeys & MText(CStr(outNames(i)))
            If CStr(keyCols(i)) <> CStr(outNames(i)) Then
                If Len(renames) > 0 Then renames = renames & ", "
                renames = renames & "{" & MText(CStr(keyCols(i))) & ", " & MText(CStr(outNames(i))) & "}"
            End If
        End If
    Next i
    If Len(amtCol) > 0 Then
        If Len(picks) > 0 Then picks = picks & ", "
        picks = picks & MText(amtCol)
        If Len(renames) > 0 Then renames = renames & ", "
        renames = renames & "{" & MText(amtCol) & ", " & MText("__Amount") & "}"
    End If

    AddLine m, "let"
    AddLine m, "    Source = " & sourceQuery & ","
    AddLine m, "    Picked = Table.SelectColumns(Source, {" & picks & "}),"
    prevStep = "Picked"
    If Len(dayCol) > 0 Then
        AddLine m, "    Dayed = Table.TransformColumns(" & prevStep & ", {{" & MText(dayCol) & _
            ", each if _ is datetime then DateTime.Date(_) else _}}),"
        prevStep = "Dayed"
    End If
    If Len(renames) > 0 Then
        AddLine m, "    Renamed = Table.RenameColumns(" & prevStep & ", {" & renames & "}),"
        prevStep = "Renamed"
    End If
    If Len(amtCol) > 0 Then
        AddLine m, "    Result = Table.Group(" & prevStep & ", {" & groupKeys & "}, {"
        AddLine m, MTpl("        {`Transaction Amount`, each List.Sum(List.Select([__Amount], each _ is number)), type nullable number},")
        AddLine m, MTpl("        {`Transaction Count`, each List.Count(List.Select([__Amount], each _ <> null and _ <> ``)), Int64.Type}})")
    Else
        AddLine m, "    Result = Table.Group(" & prevStep & ", {" & groupKeys & "}, {{" & MText("Transaction Count") & _
            ", each Table.RowCount(_), Int64.Type}})"
    End If
    AddLine m, "in"
    AddLine m, "    Result"
    BuildTotalsM = m
End Function

' Sheet7's figures over every row of TrxDeDupe, as one row:
' TxnCount, TxnSum, FirstDate, LastDate, Accounts, MinPositive,
' MaxPositive, CrSum, DrSum, CrCount, DrCount - each worked out the way its
' Sheet7 formula does it (numbers only for COUNT/SUM, dates only for
' MIN/MAX, CR/DR matched without regard to case as = and COUNTIF do,
' accounts distinct in first-seen order as UNIQUE gives them). Sheet7's
' account list only read the first 9,999 rows; this reads all of them.
Private Function BuildStatsM(ByVal colAmt As String, ByVal colDate As String, _
    ByVal colDrCr As String, ByVal colAcct As String) As String
    Dim m As String, picks As String, c As Variant, amtList As String

    For Each c In Array(colAmt, colDate, colDrCr, colAcct)
        If Len(c) > 0 Then
            If InStr(picks, MText(CStr(c))) = 0 Then
                If Len(picks) > 0 Then picks = picks & ", "
                picks = picks & MText(CStr(c))
            End If
        End If
    Next c

    AddLine m, "let"
    AddLine m, "    Source = TrxDeDupe,"
    AddLine m, "    Cols = Table.Buffer(Table.SelectColumns(Source, {" & picks & "})),"
    If Len(colAmt) > 0 Then
        AddLine m, "    Amt = Table.Column(Cols, " & MText(colAmt) & "),"
    Else
        AddLine m, "    Amt = {},"
    End If
    AddLine m, "    Nums = List.Select(Amt, each _ is number),"
    AddLine m, "    Positives = List.Select(Nums, each _ > 0),"
    If Len(colDate) > 0 Then
        AddLine m, "    Dates = List.Transform(List.Select(Table.Column(Cols, " & MText(colDate) & _
            "), each _ is date or _ is datetime), each if _ is datetime then DateTime.Date(_) else _),"
    Else
        AddLine m, "    Dates = {},"
    End If
    AddLine m, MTpl("    DateText = (d as nullable date) as text => if d = null then `01/00/1900` else Date.ToText(d, `MM/dd/yyyy`, `en-US`),")
    AddLine m, "    IsSide = (v as any, side as text) as logical => v is text and Text.Upper(v) = side,"
    If Len(colDrCr) > 0 Then
        AddLine m, "    SideRows = (side as text) as table => Table.SelectRows(Cols, each IsSide(Record.Field(_, " & MText(colDrCr) & "), side)),"
    Else
        AddLine m, "    SideRows = (side as text) as table => Table.FirstN(Cols, 0),"
    End If
    If Len(colAmt) > 0 Then
        AddLine m, "    SideSum = (side as text) => List.Sum(List.Select(Table.Column(SideRows(side), " & MText(colAmt) & "), each _ is number)),"
    Else
        AddLine m, "    SideSum = (side as text) => null,"
    End If
    AddLine m, "    SideCount = (side as text) => Table.RowCount(SideRows(side)),"
    If Len(colAcct) > 0 Then
        AddLine m, MTpl("    Accounts = List.Distinct(List.Transform(List.Select(Table.Column(Cols, ") & MText(colAcct) & _
            MTpl("), each _ <> null and _ <> ``), each Text.From(_)), Comparer.OrdinalIgnoreCase),")
    Else
        AddLine m, "    Accounts = {},"
    End If
    AddLine m, "    Result = #table("
    AddLine m, MTpl("        {`TxnCount`, `TxnSum`, `FirstDate`, `LastDate`, `Accounts`, `MinPositive`, `MaxPositive`, `CrSum`, `DrSum`, `CrCount`, `DrCount`},")
    AddLine m, "        {{List.Count(Nums), List.Sum(Nums), DateText(List.Min(Dates)), DateText(List.Max(Dates)),"
    AddLine m, MTpl("          Text.Start(Text.Combine(Accounts, `,`), 32000), List.Min(Positives), List.Max(Positives),")
    AddLine m, MTpl("          SideSum(`CR`), SideSum(`DR`), SideCount(`CR`), SideCount(`DR`)}})")
    AddLine m, "in"
    AddLine m, "    Result"
    BuildStatsM = m
End Function
