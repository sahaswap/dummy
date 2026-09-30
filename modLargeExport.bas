Attribute VB_Name = "modLargeExport"
Option Explicit

' ==========================================================
' LARGE-FILE EXPORT - a separate copy of Export Trx File
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
' LEGACY reads each source file once - Excel files in Excel, CSV files
' straight from disk (no row limit) - and does everything in that one
' pass - see LegacyExport. If a run stops after that pass,
' Consolidated_AML_Workflow_Large_Resume finishes it without re-reading. Differences from Module9's Legacy:
'   - Raw Transactions is left out (the files are in the Transaction
'     Files folder). CP Selection and DeDupe are saved as files of their
'     own - "<ECM>_<Alert>_DeDupe.xlsb", or "(part 1)", "(part 2)"... of
'     up to RAW_ROWS_PER_SHEET rows each - beside the main file (.xlsb:
'     see LIST_FILE_EXT).
'   - The main file holds the four pivots, built on two totals sheets
'     (Scenario Totals, Daily Totals) added up in that pass; a pivot adds
'     them back up, so every figure matches a pivot on the rows.
'   - When DeDupe is too big for ConsolidatedData, ConsolidatedData gets
'     only the rows that carry an alert, and Sheet7's narrative figures
'     come from that pass over every transaction - see
'     PointSheet7AtLargeStats.
'   - Counterparty is a value, not a formula; later files are matched to
'     the first one by column NAME, not position.
'
' EN and PIVOT use Power Query: it combines the folder, does the date /
' amount / Counterparty cleanup and filters each output sheet, which is
' loaded straight into its export workbook; the queries are deleted again
' before saving. Differences from Module9's output there:
'   - Counterparty is a value, not a formula.
'   - Numbers a source file holds as text stay text (Module9's value
'     rewrite turned them into numbers). Amounts are still converted.
'   - Files are matched up by column NAME when combined.
'   - A file's "Transaction ID" header must be in its first 100 rows.
'   - The Pivot Analysis export, which is saved into the same \Pivot
'     folder it reads from, is not read back in as data on the next run.
'   - EN's Raw Transactions carries on onto "Raw Transactions (2)",
'     "(3)"... when it has more rows than one sheet holds; any other sheet
'     that would need more rows than Excel allows stops the export with a
'     message instead of being cut short.
' ==========================================================

' TidyDataSheet: sheets up to this many rows get the full column + wrapped
' row fit; bigger ones fit columns to the first TIDY_SAMPLE_ROWS rows only.
Private Const TIDY_FULL_FIT_ROWS As Long = 50000
Private Const TIDY_SAMPLE_ROWS As Long = 2000

' Rows per sheet when a sheet's rows have to be spread over several (a
' sheet holds 1,048,575 below its header).
Private Const RAW_ROWS_PER_SHEET As Long = 1000000

' The step the export is on - shown on the status bar while it runs and
' named in the error message if it fails.
Private m_stage As String

' Export workbooks created this run and not yet saved.
Private m_unsaved As Collection

' Very hidden sheet in this workbook holding Sheet7's narrative figures for
' a case too big for ConsolidatedData (B1 = that case's ECM ID).
Private Const STATS_SHEET As String = "_LargeCaseStats"

' LegacyExport: rows read from a source file, and written to a temporary
' file, per go (WRITE_BLOCK_ROWS divides RAW_ROWS_PER_SHEET, so a block
' never straddles two files).
Private Const READ_BLOCK_ROWS As Long = 20000
Private Const WRITE_BLOCK_ROWS As Long = 20000
' ReadCsvSource: bytes read from a CSV per go, and how far down the file
' it looks for the "Transaction ID" header line.
Private Const CSV_CHUNK_BYTES As Long = 16777216
Private Const CSV_HEADER_SEARCH_LINES As Long = 1000

' The CP Selection / DeDupe files: .xlsb (Excel's binary format) saves
' several times faster than .xlsx and is about half the size, and opens in
' Excel the same way (File > Save As turns one into .xlsx if needed). Set
' to ".xlsx" and 51 to go back to .xlsx.
Private Const LIST_FILE_EXT As String = ".xlsb"
Private Const LIST_FILE_FORMAT As Long = 50

' Windows' own file reading, for CSV files: VBA's Open/Get can't go past
' 2 GB into a file, these have no limit.
Private Const GENERIC_READ As Long = &H80000000
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Private Const INVALID_HANDLE_VALUE As Long = -1
#If VBA7 Then
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, _
    ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, _
    ByRef Source As Any, ByVal Length As LongPtr)
Private m_csvHandle As LongPtr          ' the CSV being read (0 = none), closed on failure
#Else
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, _
    ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, _
    ByRef Source As Any, ByVal Length As Long)
Private m_csvHandle As Long
#End If
' Dictionaries per duplicate check or total - see NewShards.
Private Const SHARDS As Long = 1024
' Raised after a message has already been shown, so CancelHandler stays quiet.
Private Const ERR_REPORTED As Long = vbObjectError + 999

' One list (CP Selection, DeDupe, alerted rows) being written to temporary
' files - see TempOpen.
Private Type TempWriter
    BaseName As String
    FileNo As Integer
    Paths As Collection
    Lines() As String
    Fill As Long
    PartRows As Long
    Total As Long
End Type

' The Legacy export's columns, set from the first file's header (or, when
' resuming, a temporary file's): how many, their names and header line,
' the date column, the Counterparty column (0 = none) and which columns
' import as text; m_fields is a reusable row buffer.
Private m_outCols As Long
Private m_headerNames() As String
Private m_headerLine As String
Private m_dateCol As Long
Private m_cpTextCol As Long
Private m_textCol() As Boolean
Private m_fields() As String

' The Legacy export's scratch folder for its temporary files ("" = none),
' and whether a failure should keep them (True once they are all written,
' so Consolidated_AML_Workflow_Large_Resume can finish the export).
Private m_tempFolder As String
Private m_keepTemp As Boolean

' Pass 1's state, shared by the .xlsx and .csv readers - see LegacyExport
' and SetMasterColumns: the first file's header and the columns found in
' it, the duplicate checks, the three temporary lists and the row count.
Private m_haveMaster As Boolean, m_master() As String, m_nMaster As Long
Private m_p1Date As Long, m_p1Amt As Long, m_p1Dr As Long, m_p1Ben As Long, m_p1Orig As Long
Private m_p1Trans As Long, m_p1Alert As Long
Private m_hasCp As Boolean, m_dedupeCP As Boolean, m_dedupeDD As Boolean
Private m_cpSeen As Variant, m_ddSeen As Variant, m_alertIds As Object
Private m_cpW As TempWriter, m_ddW As TempWriter, m_alW As TempWriter
Private m_rowVals() As Variant, m_rowsRead As Double, m_extraCols As Long
' ProcessCsvFields: which CSV field each export column comes from, and the
' row being built.
Private m_csvMap() As Long, m_csvOut() As String
' How long pass 1 took, for the completion message (-1 = resumed, not read).
Private m_readSecs As Double

' Pass 2's totals and Sheet7 figures - see ResetTotals.
Private m_iAlert As Long, m_iDr As Long, m_iAmt As Long, m_iAcct As Long
Private m_scnIdx As Variant, m_scnN As Long, m_scnCap As Long
Private m_scnAlert() As Variant, m_scnDr() As Variant, m_scnCp() As Variant
Private m_scnSum() As Double, m_scnCnt() As Long
Private m_dayIdx As Variant, m_dayN As Long, m_dayCap As Long
Private m_dayDate() As Variant, m_dayDr() As Variant, m_daySum() As Double, m_dayCnt() As Long
Private m_stCount As Double, m_stSum As Double, m_stHaveDate As Boolean, m_stFirst As Date, m_stLast As Date
Private m_stHavePos As Boolean, m_stMinPos As Double, m_stMaxPos As Double
Private m_crSum As Double, m_drSum As Double, m_crCnt As Double, m_drCnt As Double
Private m_acctSeen As Variant, m_acctText As String

Sub Consolidated_AML_Workflow_Large()

' ==========================================
' LARGE-FILE EXPORT (POWER QUERY)
' ==========================================
Application.EnableCancelKey = xlErrorHandler
On Error GoTo CancelHandler
m_stage = "starting the export"
Set m_unsaved = New Collection
m_keepTemp = False

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

' Legacy also reads CSV files (see ReadCsvSource); EN and Pivot don't.
For Each objFile In objFolder.Files
    If exportMode <> "EN" And exportMode <> "PIVOT" Then
        fileFound = IsLegacySourceFile(objFile.Name)
    Else
        fileFound = (InStr(1, objFile.Name, ".xls", vbTextCompare) > 0) And (Left(objFile.Name, 2) <> "~$") And (objFile.Name <> ThisWorkbook.Name)
    End If
    If fileFound Then Exit For
Next objFile

If Not fileFound Then
    If exportMode <> "EN" And exportMode <> "PIVOT" Then
        MsgBox "No Excel or CSV files found in the target folder!", vbExclamation, "Folder is Empty"
    Else
        MsgBox "No Excel files found in the target folder!", vbExclamation, "Folder is Empty"
    End If
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
' 4. TWO-LAYER EXPORT & DEDUPE (LEGACY) - one pass, no Power Query
' ==========================================
If exportMode <> "EN" Then
    LegacyExport wsHome, wsRealCD, ecmID, AlertID, folderPath, saveFolderPath, origCalc
    Exit Sub
End If

' ==========================================
' 2-3. EN: COMBINE + CLEANUP (Power Query)
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
Close                       ' any temporary file LegacyExport had open
CloseCsvHandle
' Once the source files are all read their rows are kept, for the Resume
' macro; before that there is nothing worth keeping.
If Not m_keepTemp Then RemoveTempFolder
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
ElseIf savedErrNum <> 0 And savedErrNum <> ERR_REPORTED Then
    If m_keepTemp Then
        MsgBox "An unexpected error occurred while " & m_stage & ":" & vbCrLf & vbCrLf & _
            "Error " & savedErrNum & ": " & savedErrDesc & vbCrLf & vbCrLf & _
            "The source files were already read and those rows are kept - run " & _
            "Consolidated_AML_Workflow_Large_Resume to finish from here.", vbCritical, "Large Export Error"
    Else
        MsgBox "An unexpected error occurred while " & m_stage & ":" & vbCrLf & vbCrLf & _
            "Error " & savedErrNum & ": " & savedErrDesc, vbCritical, "Large Export Error"
    End If
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
' those figures from LegacyExport's pass instead, stored on the very
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

' stats(1 To 11), from LegacyExport's pass: count, sum, first date, last
' date, accounts, smallest and largest amount over 0, CR sum, DR sum, CR
' count, DR count. Stored exactly as each Sheet7 formula shows it (its TEXT()
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

' ==========================================================
' LEGACY EXPORT - every file read once; lists saved as separate files
' ==========================================================
' Pass 1 (LegacyExport) reads each source file in turn - an Excel file
' opened in Excel (much faster than Power Query's .xlsx reader), a CSV
' straight from disk (ReadCsvSource) - READ_BLOCK_ROWS rows at a time, and
' deals with every row in that one pass (ProcessBlock):
'   - cleaned as Module9 cleans it: text dates read month/day/year, text
'     amounts made numbers, Counterparty = Beneficiary Name on DR rows and
'     Originator Name otherwise;
'   - kept for CP Selection unless its Transaction ID + Alert Information
'     pair was already seen, and for DeDupe unless its Transaction ID was
'     already seen or it has no date - Module9's two RemoveDuplicates
'     passes and its blank-date delete, in the same order, the first row
'     kept each time;
'   - kept, if it carries an alert, for ConsolidatedData when DeDupe is too
'     big for it.
' The rows kept go to temporary text files on disk, not into Excel.
'
' Pass 2 (FinishLegacyExport) turns each temporary file into its own .xlsx
' of up to RAW_ROWS_PER_SHEET rows, saved and closed before the next one,
' adding up the pivots' totals and Sheet7's figures from each as it goes,
' then builds the main file.
'
' Why the lists are separate files: with 30+ lakh rows, CP Selection and
' DeDupe come to 50-60 lakh rows. One workbook holding them, however many
' sheets it's split over, keeps them all in memory at once, and on the VDI
' Excel ran out of memory part-way through (it failed opening the second
' source file, reporting it as damaged, though it opens fine on its own).
' This way Excel holds one source file, or one output file, at a time.
'
' Every file is built in the local scratch folder, which OneDrive doesn't
' watch, and only moved into the case folder once all of them exist, so
' OneDrive syncs the finished set in one go.
'
' Once pass 1 is done its temporary files are kept if anything later
' fails, and Consolidated_AML_Workflow_Large_Resume finishes the export
' from them without reading the source files again.
'
' Later files are matched to the first file's header by column name;
' columns the first file doesn't have are left out and counted in the
' completion message. Raw Transactions is left out (the files themselves
' are in the Transaction Files folder).
' ==========================================================
Private Sub LegacyExport(ByVal wsHome As Worksheet, ByVal wsRealCD As Worksheet, _
    ByVal ecmID As String, ByVal AlertID As String, ByVal folderPath As String, _
    ByVal saveFolderPath As String, ByVal origCalc As XlCalculation)

    Dim FSO As Object, f As Object, files As Collection, item As Variant
    Dim fileNo As Long, filesRead As Long, skippedFiles As String, summary As String, found As Boolean

    ' Module9's file choice, plus .csv files.
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set files = New Collection
    For Each f In FSO.GetFolder(folderPath).files
        If IsLegacySourceFile(f.Name) Then files.Add f.Path
    Next f

    ' Scratch folder for the temporary files (local, not synced by OneDrive).
    m_tempFolder = Environ$("TEMP") & Application.PathSeparator & "LargeExport_" & Format$(Now, "yyyymmdd_hhnnss")
    MkDir m_tempFolder

    m_readSecs = Timer
    m_haveMaster = False
    m_rowsRead = 0
    m_extraCols = 0
    m_cpSeen = NewShards()
    m_ddSeen = NewShards()
    Set m_alertIds = CreateObject("Scripting.Dictionary")

    For Each item In files
        fileNo = fileNo + 1
        If LCase$(Right$(CStr(item), 4)) = ".csv" Then
            found = ReadCsvSource(CStr(item), fileNo, files.count)
        Else
            found = ReadXlsxSource(CStr(item), fileNo, files.count)
        End If
        If found Then
            filesRead = filesRead + 1
        Else
            skippedFiles = skippedFiles & vbCrLf & "   " & FileNameOf(CStr(item))
        End If
    Next item

    If Not m_haveMaster Then
        MsgBox "None of the Excel or CSV files in:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
            "has a 'Transaction ID' header, so there is nothing to export.", vbCritical, "No Transaction Data"
        Err.Raise ERR_REPORTED
    End If

    TempClose m_cpW
    TempClose m_ddW
    TempClose m_alW
    ' The duplicate checks are done with; free their memory before pass 2.
    m_cpSeen = Empty
    m_ddSeen = Empty
    Set m_alertIds = Nothing

    ' From here on a failure keeps the temporary files for the Resume macro.
    m_keepTemp = True
    m_readSecs = Timer - m_readSecs
    If m_readSecs < 0 Then m_readSecs = m_readSecs + 86400

    summary = Format$(m_rowsRead, "#,##0") & " rows read from " & filesRead & " file(s)."
    If Len(skippedFiles) > 0 Then
        summary = summary & vbCrLf & vbCrLf & "Skipped - no 'Transaction ID' header:" & skippedFiles
    End If
    If m_extraCols > 0 Then
        summary = summary & vbCrLf & vbCrLf & m_extraCols & " column(s) in later files aren't in the " & _
            "first file's header and were left out."
    End If

    FinishLegacyExport wsHome, wsRealCD, ecmID, AlertID, saveFolderPath, origCalc, _
        m_cpW.Paths, m_ddW.Paths, m_alW.Paths, summary
End Sub

' Excel files as Module9 picks them (".xls" anywhere in the name), and CSV
' files; never Excel's "~$" lock files or this workbook.
Private Function IsLegacySourceFile(ByVal fileName As String) As Boolean
    If Left$(fileName, 2) = "~$" Or fileName = ThisWorkbook.Name Then Exit Function
    IsLegacySourceFile = (InStr(1, fileName, ".xls", vbTextCompare) > 0) Or _
        (LCase$(Right$(fileName, 4)) = ".csv")
End Function

' The first file's header sets the export's columns (the header line, the
' columns Module9 finds by name, and whether Counterparty is added), and
' opens the three temporary lists.
Private Sub SetMasterColumns(ByRef fileHdr() As String)
    Dim j As Long
    m_master = fileHdr
    m_nMaster = UBound(m_master)
    m_p1Date = FindHeaderIndex(m_master, "Transaction Date", False)
    m_p1Amt = FindHeaderIndex(m_master, "Transaction Amount", False)
    m_p1Dr = FindHeaderIndex(m_master, "Dr Cr", False)
    m_p1Ben = FindHeaderIndex(m_master, "Beneficiary Name", False)
    m_p1Orig = FindHeaderIndex(m_master, "Originator Name", False)
    m_p1Trans = FindHeaderIndex(m_master, "Transaction ID", False)
    m_p1Alert = FindHeaderIndex(m_master, "Alert Information", False)
    m_hasCp = (m_p1Dr > 0 And m_p1Ben > 0 And m_p1Orig > 0)
    m_dedupeCP = (m_p1Trans > 0 And m_p1Alert > 0)
    m_dedupeDD = (m_p1Trans > 0)

    m_outCols = m_nMaster
    If m_hasCp Then m_outCols = m_nMaster + 1
    ReDim m_headerNames(1 To m_outCols)
    ReDim m_fields(1 To m_outCols)
    ReDim m_textCol(1 To m_outCols)
    For j = 1 To m_nMaster
        m_headerNames(j) = CleanFieldText(m_master(j))
    Next j
    m_cpTextCol = 0
    If m_hasCp Then
        m_headerNames(m_outCols) = "Counterparty"
        m_cpTextCol = m_outCols
    End If
    m_headerLine = Join(m_headerNames, vbTab)
    m_dateCol = m_p1Date
    ReDim m_rowVals(1 To m_outCols)
    ReDim m_csvOut(1 To m_outCols)

    TempOpen m_cpW, "CP Selection"
    TempOpen m_ddW, "DeDupe"
    TempOpen m_alW, "Alerted Rows"
    m_haveMaster = True
End Sub

' Pass 1 for a block of source rows - the same for .xlsx and .csv files.
' blk(r, c) is row r, column c of the file (from its Transaction ID
' column); colMap says which file column each export column comes from.
Private Sub ProcessBlock(ByRef blk As Variant, ByVal nBlk As Long, ByRef colMap() As Long)
    Dim r As Long, j As Long, c As Long, v As Variant, isBlank As Boolean, isDrRow As Boolean
    Dim k As String, rowText As String

    For r = 1 To nBlk
        ' Row in the export's column order; wholly empty rows skipped.
        isBlank = True
        For j = 1 To m_nMaster
            c = colMap(j)
            If c > 0 Then v = blk(r, c) Else v = Empty
            If isBlank Then isBlank = IsBlankValue(v)
            m_rowVals(j) = v
        Next j
        If isBlank Then GoTo NextRow

        ' Module9's step 3 cleanup
        If m_p1Date > 0 Then m_rowVals(m_p1Date) = CleanDate(m_rowVals(m_p1Date))
        If m_p1Amt > 0 Then m_rowVals(m_p1Amt) = CleanAmount(m_rowVals(m_p1Amt))
        If m_hasCp Then
            ' =IF(DrCr="DR", IF(Ben="","",Ben), IF(Orig="","",Orig))
            isDrRow = False
            v = m_rowVals(m_p1Dr)
            If VarType(v) = vbString Then isDrRow = (StrComp(v, "DR", vbTextCompare) = 0)
            If isDrRow Then v = m_rowVals(m_p1Ben) Else v = m_rowVals(m_p1Orig)
            If IsBlankValue(v) Then v = ""
            m_rowVals(m_outCols) = v
        End If

        ' ---- CP Selection: first row per Transaction ID + Alert Information.
        ' The alert text is swapped for a short number first, so the lakhs of
        ' keys held here stay small.
        If m_dedupeCP Then
            k = KeyText(m_rowVals(m_p1Alert))
            If Not m_alertIds.Exists(k) Then m_alertIds.Add k, m_alertIds.count + 1
            If AlreadySeen(m_cpSeen, m_alertIds(k) & vbTab & KeyText(m_rowVals(m_p1Trans))) Then GoTo NextRow
        End If
        rowText = RowLine(m_rowVals)
        TempAdd m_cpW, rowText

        ' ---- DeDupe: first row per Transaction ID, then rows with a date
        If m_dedupeDD Then
            If AlreadySeen(m_ddSeen, KeyText(m_rowVals(m_p1Trans))) Then GoTo NextRow
        End If
        If m_p1Date > 0 Then
            If IsBlankValue(m_rowVals(m_p1Date)) Then GoTo NextRow
        End If
        TempAdd m_ddW, rowText

        ' Rows that carry an alert, for ConsolidatedData if DeDupe won't fit
        If m_p1Alert > 0 Then
            If Not IsBlankValue(m_rowVals(m_p1Alert)) Then TempAdd m_alW, rowText
        End If
NextRow:
    Next r
    m_rowsRead = m_rowsRead + nBlk
End Sub

Private Sub ShowReadProgress(ByVal fileNo As Long, ByVal fileCount As Long)
    SetStage "file " & fileNo & " of " & fileCount & ": " & Format$(m_rowsRead, "#,##0") & _
        " rows read, " & Format$(m_ddW.Total + m_ddW.Fill, "#,##0") & " unique transactions"
    DoEvents
End Sub

' ==========================================================
' ReadXlsxSource - an Excel source file, opened in Excel
' ==========================================================
' Its first sheet, from the "Transaction ID" header cell (whole cell, any
' case - Module9's Find) across and down, read READ_BLOCK_ROWS rows at a
' time. False if the sheet has no such header (the file is skipped).
' ==========================================================
Private Function ReadXlsxSource(ByVal path As String, ByVal fileNo As Long, ByVal fileCount As Long) As Boolean
    Dim wbSrc As Workbook, wsSrc As Worksheet, hc As Range
    Dim hRow As Long, hCol As Long, lastR As Long, lastC As Long, startR As Long, endR As Long
    Dim fileHdr() As String, colMap() As Long, blk As Variant

    SetStage "opening file " & fileNo & " of " & fileCount & " (" & FileNameOf(path) & ")"
    DoEvents
    Set wbSrc = Workbooks.Open(fileName:=path, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    m_unsaved.Add wbSrc

    If TypeName(wbSrc.Sheets(1)) = "Worksheet" Then
        Set wsSrc = wbSrc.Sheets(1)
        Set hc = wsSrc.Cells.Find(What:="Transaction ID", LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False)
    End If

    If Not hc Is Nothing Then
        ReadXlsxSource = True
        hRow = hc.row
        hCol = hc.Column
        lastC = wsSrc.Cells(hRow, wsSrc.Columns.count).End(xlToLeft).Column
        If lastC < hCol Then lastC = hCol
        lastR = LastDataRow(wsSrc)
        fileHdr = RowTexts(wsSrc, hRow, hCol, lastC)
        If Not m_haveMaster Then SetMasterColumns fileHdr
        colMap = MapColumns(m_master, fileHdr, m_extraCols)

        For startR = hRow + 1 To lastR Step READ_BLOCK_ROWS
            endR = startR + READ_BLOCK_ROWS - 1
            If endR > lastR Then endR = lastR
            blk = wsSrc.Range(wsSrc.Cells(startR, hCol), wsSrc.Cells(endR, lastC)).Value
            If Not IsArray(blk) Then blk = OneCellArray(blk)
            ProcessBlock blk, endR - startR + 1, colMap
            ShowReadProgress fileNo, fileCount
        Next startR
    End If

    MarkSaved wbSrc
    wbSrc.Close SaveChanges:=False
End Function

' ==========================================================
' ReadCsvSource - a CSV source file, read straight from disk
' ==========================================================
' Not opened in Excel: Excel loads only the first 1,048,576 rows of a CSV
' (and with alerts off, silently), and CSV exports this size are often
' longer than that. Instead the file is read CSV_CHUNK_BYTES at a time
' through Windows' own file reading (no size limit - VBA's Open/Get stop at
' 2 GB), each chunk used up to its last line break so no row is cut in
' two, the rest carried to the front of the next chunk.
'   - Encoding: UTF-8 (with or without its byte-order mark) or UTF-16 by
'     its mark; a file that isn't valid UTF-8 is read as Windows ANSI.
'   - Separator: whichever of , ; tab | the header line uses.
'   - Quoted fields ("Smith, John", "say ""hi""") are handled, including
'     ones with a line break inside.
'   - The header is the first line with a "Transaction ID" field (any
'     case) within the first CSV_HEADER_SEARCH_LINES lines; columns start
'     at that field, as in an Excel file.
' Each row goes through ProcessCsvFields, which does ProcessBlock's work
' on the text directly. False if no header was found (the file is skipped).
' ==========================================================
Private Function ReadCsvSource(ByVal path As String, ByVal fileNo As Long, ByVal fileCount As Long) As Boolean
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    Dim buf() As Byte, part() As Byte, have As Long, got As Long, total As Long
    Dim startAt As Long, useEnd As Long, atEof As Boolean, firstChunk As Boolean, utf16 As Boolean
    Dim charset As String, text As String, lines() As String, i As Long, lastI As Long
    Dim s As String, pending As String, fields() As String, delim As String, cleanNeeded As Boolean
    Dim haveHeader As Boolean, scanned As Long, hIdx As Long, nHdr As Long, lastField As Long
    Dim fileHdr() As String, colMap() As Long, j As Long, sinceProgress As Long

    SetStage "reading file " & fileNo & " of " & fileCount & " (" & FileNameOf(path) & ")"
    DoEvents
    h = CreateFileW(StrPtr(path), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, _
        OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If h = INVALID_HANDLE_VALUE Then
        Err.Raise vbObjectError + 1007, "modLargeExport", "Couldn't open " & FileNameOf(path) & _
            " - is it open in another program?"
    End If
    m_csvHandle = h

    charset = "utf-8"
    firstChunk = True
    ReDim buf(0 To CSV_CHUNK_BYTES - 1)
    Do
        got = 0
        If ReadFile(h, buf(have), CSV_CHUNK_BYTES - have, got, 0) = 0 Then
            Err.Raise vbObjectError + 1009, "modLargeExport", "Reading " & FileNameOf(path) & " failed part-way."
        End If
        total = have + got
        If total = 0 Then Exit Do
        atEof = (total < CSV_CHUNK_BYTES)       ' a short read is the end of the file

        ' Byte-order mark, at the very start: UTF-16 LE (FF FE) or UTF-8 (EF BB BF).
        startAt = 0
        If firstChunk And total >= 2 Then
            If buf(0) = &HFF And buf(1) = &HFE Then
                utf16 = True
                startAt = 2
            ElseIf total >= 3 Then
                If buf(0) = &HEF And buf(1) = &HBB And buf(2) = &HBF Then startAt = 3
            End If
        End If

        If atEof Then
            useEnd = total
        Else
            useEnd = LastLineBreakEnd(buf, total, utf16)
            If useEnd <= startAt Then
                Err.Raise vbObjectError + 1008, "modLargeExport", FileNameOf(path) & _
                    " has a line longer than " & Format$(CSV_CHUNK_BYTES / 1048576, "0") & " MB."
            End If
        End If

        text = ""
        If useEnd > startAt Then
            ReDim part(0 To useEnd - startAt - 1)
            CopyMemory part(0), buf(startAt), useEnd - startAt
            If utf16 Then
                text = part                         ' UTF-16 bytes are a VBA string as they are
            Else
                text = DecodeBytes(part, charset)
                If firstChunk Then
                    If InStr(text, ChrW(&HFFFD)) > 0 Then
                        charset = "windows-1252"    ' not valid UTF-8: Windows ANSI
                        text = DecodeBytes(part, charset)
                    End If
                End If
            End If
        End If
        firstChunk = False

        ' What's past the last line break goes to the front of the next chunk.
        have = total - useEnd
        If have > 0 Then CopyMemory buf(0), buf(useEnd), have

        lines = Split(text, vbLf)
        lastI = UBound(lines)
        If lastI >= 0 Then
            If Len(lines(lastI)) = 0 Then lastI = lastI - 1   ' the chunk ends with its line break
        End If

        For i = 0 To lastI
            s = lines(i)
            If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
            If Len(pending) > 0 Then
                s = pending & vbLf & s
                pending = ""
            End If

            If Not haveHeader Then
                scanned = scanned + 1
                If FindCsvHeader(s, delim, fields, hIdx) Then
                    lastField = UBound(fields)
                    Do While lastField > hIdx And Len(Trim$(fields(lastField))) = 0
                        lastField = lastField - 1
                    Loop
                    nHdr = lastField - hIdx + 1
                    ReDim fileHdr(1 To nHdr)
                    For j = 1 To nHdr
                        fileHdr(j) = Trim$(fields(hIdx + j - 1))
                    Next j
                    If Not m_haveMaster Then SetMasterColumns fileHdr
                    colMap = MapColumns(m_master, fileHdr, m_extraCols)
                    ' Export column j comes from field m_csvMap(j) of the line (-1 = none).
                    ReDim m_csvMap(1 To m_nMaster)
                    For j = 1 To m_nMaster
                        If colMap(j) > 0 Then m_csvMap(j) = hIdx + colMap(j) - 1 Else m_csvMap(j) = -1
                    Next j
                    haveHeader = True
                ElseIf scanned >= CSV_HEADER_SEARCH_LINES Then
                    GoTo CsvDone                   ' no header near the top: skip the file
                End If
            Else
                If InStr(s, """") > 0 Then
                    If Not ParseCsvLine(s, delim, fields) Then
                        pending = s                ' a quoted field runs onto the next line
                        GoTo NextLine
                    End If
                Else
                    fields = Split(s, delim)
                End If
                ' A tab or line break inside a field would split the row in the
                ' temporary file; only lines that have one pay for the clean-up.
                cleanNeeded = (InStr(s, vbLf) > 0) Or (InStr(s, vbCr) > 0)
                If Not cleanNeeded And delim <> vbTab Then cleanNeeded = (InStr(s, vbTab) > 0)
                ProcessCsvFields fields, cleanNeeded
                sinceProgress = sinceProgress + 1
                If sinceProgress = READ_BLOCK_ROWS Then
                    sinceProgress = 0
                    ShowReadProgress fileNo, fileCount
                End If
            End If
NextLine:
        Next i
        DoEvents
        If atEof Then Exit Do
    Loop

    ' A last row whose quotes never closed: taken as it stands.
    If haveHeader And Len(pending) > 0 Then
        fields = Split(Replace(pending, """", ""), delim)
        ProcessCsvFields fields, True
    End If
    If haveHeader Then ShowReadProgress fileNo, fileCount

CsvDone:
    CloseCsvHandle
    ReadCsvSource = haveHeader
End Function

Private Sub CloseCsvHandle()
    If m_csvHandle <> 0 Then
        CloseHandle m_csvHandle
        m_csvHandle = 0
    End If
End Sub

' ==========================================================
' ProcessCsvFields - ProcessBlock's work for one CSV row, on the text
' ==========================================================
' A CSV row is text already, so instead of copying every field into a
' Variant and writing it back out (ProcessBlock), the fields go straight
' to the output line; only the date, the amount and Counterparty are
' worked on. The line written, and every duplicate and blank check, are
' the same as ProcessBlock would make of the same row.
' ==========================================================
Private Sub ProcessCsvFields(ByRef fields() As String, ByVal cleanNeeded As Boolean)
    Dim j As Long, idx As Long, ub As Long, isBlank As Boolean, rowText As String, k As String

    m_rowsRead = m_rowsRead + 1
    ub = UBound(fields)
    isBlank = True
    For j = 1 To m_nMaster
        idx = m_csvMap(j)
        If idx >= 0 And idx <= ub Then
            m_csvOut(j) = fields(idx)
            If isBlank Then isBlank = (Len(m_csvOut(j)) = 0)
        Else
            m_csvOut(j) = vbNullString
        End If
    Next j
    If isBlank Then Exit Sub

    If cleanNeeded Then
        For j = 1 To m_nMaster
            m_csvOut(j) = CleanFieldText(m_csvOut(j))
        Next j
    End If

    ' Module9's step 3 cleanup
    If m_p1Date > 0 Then m_csvOut(m_p1Date) = CsvDateText(m_csvOut(m_p1Date))
    If m_p1Amt > 0 Then m_csvOut(m_p1Amt) = CsvAmountText(m_csvOut(m_p1Amt))
    If m_hasCp Then
        ' =IF(DrCr="DR", IF(Ben="","",Ben), IF(Orig="","",Orig))
        If StrComp(m_csvOut(m_p1Dr), "DR", vbTextCompare) = 0 Then
            m_csvOut(m_outCols) = m_csvOut(m_p1Ben)
        Else
            m_csvOut(m_outCols) = m_csvOut(m_p1Orig)
        End If
    End If

    ' Columns of over-long digit strings import as text (see FieldText).
    For j = 1 To m_outCols
        If Len(m_csvOut(j)) > 15 Then
            If Not m_textCol(j) Then
                If Not (m_csvOut(j) Like "*[!0-9]*") Then m_textCol(j) = True
            End If
        End If
    Next j

    ' ---- CP Selection: first row per Transaction ID + Alert Information
    If m_dedupeCP Then
        k = UCase$(m_csvOut(m_p1Alert))
        If Not m_alertIds.Exists(k) Then m_alertIds.Add k, m_alertIds.count + 1
        If AlreadySeen(m_cpSeen, m_alertIds(k) & vbTab & UCase$(m_csvOut(m_p1Trans))) Then Exit Sub
    End If
    rowText = Join(m_csvOut, vbTab)
    TempAdd m_cpW, rowText

    ' ---- DeDupe: first row per Transaction ID, then rows with a date
    If m_dedupeDD Then
        If AlreadySeen(m_ddSeen, UCase$(m_csvOut(m_p1Trans))) Then Exit Sub
    End If
    If m_p1Date > 0 Then
        If Len(m_csvOut(m_p1Date)) = 0 Then Exit Sub
    End If
    TempAdd m_ddW, rowText

    ' Rows that carry an alert, for ConsolidatedData if DeDupe won't fit
    If m_p1Alert > 0 Then
        If Len(m_csvOut(m_p1Alert)) > 0 Then TempAdd m_alW, rowText
    End If
End Sub

' A CSV date as the temporary file should hold it - what CleanDate then
' FieldText make of it. Already mm/dd/yyyy: written as it is (a valid one
' would be re-written identically, an invalid one stays the same text).
Private Function CsvDateText(ByRef s As String) As String
    Dim v As Variant
    If Len(s) = 0 Then Exit Function
    If s Like "##/##/####" Then
        CsvDateText = s
        Exit Function
    End If
    v = ParseMdyText(s)
    Select Case VarType(v)
        Case vbDate
            If v = Int(v) Then
                CsvDateText = Format$(v, "mm\/dd\/yyyy")
            Else
                CsvDateText = Format$(v, "mm\/dd\/yyyy hh:nn:ss")
            End If
        Case vbEmpty
            ' blank
        Case Else
            CsvDateText = s
    End Select
End Function

' A CSV amount as the temporary file should hold it - what CleanAmount then
' FieldText make of it. Already a plain number: written as it is (Excel
' reads "5600.00" and "5600" as the same number).
Private Function CsvAmountText(ByRef s As String) As String
    Dim v As Variant
    If Len(s) = 0 Then Exit Function
    If LooksNumeric(s) Then
        CsvAmountText = s
        Exit Function
    End If
    v = CleanAmount(s)
    If VarType(v) = vbString Then
        CsvAmountText = s
    ElseIf Not IsEmpty(v) Then
        CsvAmountText = Trim$(Str$(v))
    End If
End Function

' "4 min 12 s" for a number of seconds.
Private Function DurationText(ByVal secs As Double) As String
    If secs < 0 Then secs = secs + 86400      ' the run crossed midnight
    secs = Int(secs + 0.5)
    If secs >= 60 Then
        DurationText = Int(secs / 60) & " min " & (secs Mod 60) & " s"
    Else
        DurationText = secs & " s"
    End If
End Function

' How many bytes of buf(0..n-1) run up to and including its last line
' break (0 if none). A line break byte never falls inside a UTF-8
' character, so cutting there never splits one; for UTF-16 the break is
' the two bytes 0A 00 at an even offset.
Private Function LastLineBreakEnd(ByRef buf() As Byte, ByVal n As Long, ByVal utf16 As Boolean) As Long
    Dim k As Long
    If utf16 Then
        k = n - 2
        If k Mod 2 = 1 Then k = k - 1
        Do While k >= 0
            If buf(k) = 10 And buf(k + 1) = 0 Then
                LastLineBreakEnd = k + 2
                Exit Function
            End If
            k = k - 2
        Loop
    Else
        For k = n - 1 To 0 Step -1
            If buf(k) = 10 Then
                LastLineBreakEnd = k + 1
                Exit Function
            End If
        Next k
    End If
End Function

' Bytes to text in the given character set.
Private Function DecodeBytes(ByRef buf() As Byte, ByVal charset As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                ' binary
    stm.Open
    stm.Write buf
    stm.Position = 0
    stm.Type = 2                ' text
    stm.charset = charset
    DecodeBytes = stm.ReadText
    stm.Close
End Function

' Is line s the CSV header? Tries each common separator; on a match sets
' delim, fields (the line split by it) and hIdx (the Transaction ID field).
Private Function FindCsvHeader(ByVal s As String, ByRef delim As String, ByRef fields() As String, _
    ByRef hIdx As Long) As Boolean
    Dim d As Variant, f() As String, i As Long, ok As Boolean
    For Each d In Array(",", ";", vbTab, "|")
        If InStr(s, d) > 0 Then
            If InStr(s, """") > 0 Then
                ok = ParseCsvLine(s, CStr(d), f)
            Else
                f = Split(s, d)
                ok = True
            End If
            If ok Then
                For i = 0 To UBound(f)
                    If StrComp(Trim$(f(i)), "Transaction ID", vbTextCompare) = 0 Then
                        delim = CStr(d)
                        fields = f
                        hIdx = i
                        FindCsvHeader = True
                        Exit Function
                    End If
                Next i
            End If
        End If
    Next d
End Function

' Splits a CSV line that has quotes in it: a field in "..." may hold the
' separator, and "" inside it is one quote. False if a quoted field is
' still open at the end of the line (its line break is part of the field,
' so the next line continues it).
Private Function ParseCsvLine(ByVal s As String, ByVal delim As String, ByRef fields() As String) As Boolean
    Dim buf() As String, cnt As Long, pos As Long, n As Long, q As Long, d As Long, cur As String

    ReDim buf(0 To 63)
    n = Len(s)
    pos = 1
    Do
        If cnt > UBound(buf) Then ReDim Preserve buf(0 To UBound(buf) * 2 + 1)
        If pos > n Then
            buf(cnt) = ""                   ' empty last field (line ended with a separator)
            cnt = cnt + 1
            Exit Do
        End If
        If Mid$(s, pos, 1) = """" Then
            cur = ""
            pos = pos + 1
            Do
                q = InStr(pos, s, """")
                If q = 0 Then Exit Function           ' still inside quotes: continues on the next line
                If q < n And Mid$(s, q + 1, 1) = """" Then
                    cur = cur & Mid$(s, pos, q - pos) & """"
                    pos = q + 2
                Else
                    cur = cur & Mid$(s, pos, q - pos)
                    pos = q + 1
                    Exit Do
                End If
            Loop
            d = InStr(pos, s, delim)
            If d = 0 Then
                buf(cnt) = cur & Mid$(s, pos)
                cnt = cnt + 1
                Exit Do
            End If
            buf(cnt) = cur & Mid$(s, pos, d - pos)
            cnt = cnt + 1
            pos = d + Len(delim)
        Else
            d = InStr(pos, s, delim)
            If d = 0 Then
                buf(cnt) = Mid$(s, pos)
                cnt = cnt + 1
                Exit Do
            End If
            buf(cnt) = Mid$(s, pos, d - pos)
            cnt = cnt + 1
            pos = d + Len(delim)
        End If
    Loop
    ReDim Preserve buf(0 To cnt - 1)
    fields = buf
    ParseCsvLine = True
End Function

' ==========================================================
' RESUME - finishes a Legacy large export whose pass 1 is done
' ==========================================================
' For a run that stopped after reading the source files - stuck saving,
' Excel closed, an error - whose temporary files are still in the scratch
' folder. Finds the newest one, checks with the user, and runs pass 2 from
' it: the source files are not read again. The export is saved for the
' ECM ID / Alert ID now on Sheet1, so they must be the same case.
' Run it from Developer > Macros > Consolidated_AML_Workflow_Large_Resume.
' ==========================================================
Public Sub Consolidated_AML_Workflow_Large_Resume()
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo ResumeFailed
    m_stage = "starting to resume"
    Set m_unsaved = New Collection
    m_keepTemp = True

    Dim origCalc As XlCalculation, wsHome As Worksheet, wsRealCD As Worksheet
    Dim ecmID As String, AlertID As String, saveFolderPath As String, slash As String
    Dim folder As String, cpPaths As Collection, ddPaths As Collection, alPaths As Collection
    Dim stamp As String, stoppedAt As String

    origCalc = Application.Calculation
    slash = Application.PathSeparator

    Set wsHome = ThisWorkbook.Sheets("Sheet1")
    ecmID = Trim(wsHome.Range("J9").Value)
    AlertID = Trim(wsHome.Range("J10").Value)
    If AlertID = "" Then AlertID = "ALERT"
    If ecmID = "" Then
        MsgBox "Action Denied: ECM ID is missing in J9.", vbCritical, "Missing ID"
        Exit Sub
    End If
    saveFolderPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & slash & ecmID
    If Len(Dir(saveFolderPath, vbDirectory)) = 0 Then
        MsgBox "The case folder for ECM ID " & ecmID & " wasn't found on the Desktop.", vbCritical, "Folder Not Found"
        Exit Sub
    End If

    folder = LatestUnfinishedExport()
    If Len(folder) = 0 Then
        MsgBox "There is no unfinished large export to continue - its temporary files are " & _
            "gone (a run cancelled or failed before the rows were all read deletes them)." & vbCrLf & vbCrLf & _
            "Run Consolidated_AML_Workflow_Large again.", vbExclamation, "Nothing to Resume"
        Exit Sub
    End If
    Set cpPaths = TempPartPaths(folder, "CP Selection")
    Set ddPaths = TempPartPaths(folder, "DeDupe")
    Set alPaths = TempPartPaths(folder, "Alerted Rows")

    stamp = Mid$(folder, InStrRev(folder, "LargeExport_") + Len("LargeExport_"))
    stoppedAt = Mid$(stamp, 7, 2) & "/" & Mid$(stamp, 5, 2) & "/" & Left$(stamp, 4) & " " & _
        Mid$(stamp, 10, 2) & ":" & Mid$(stamp, 12, 2)
    If MsgBox("Continue the large export started " & stoppedAt & "?" & vbCrLf & vbCrLf & _
        "Its source files are already read: " & cpPaths.count & " CP Selection and " & ddPaths.count & _
        " DeDupe file(s) are waiting to be saved." & vbCrLf & vbCrLf & _
        "They will be saved for ECM ID " & ecmID & " / " & AlertID & " (from Sheet1) - make sure that " & _
        "is the case this export was for.", vbQuestion + vbYesNo, "Resume Large Export") <> vbYes Then Exit Sub

    ThisWorkbook.Unprotect Password:="p7ss"
    On Error Resume Next
    ThisWorkbook.Sheets("Sheet1").Unprotect Password:="p7ss"
    ThisWorkbook.Sheets("ConsolidatedData").Unprotect Password:="p7ss"
    Set wsRealCD = ThisWorkbook.Sheets("ConsolidatedData")
    On Error GoTo ResumeFailed
    If wsRealCD Is Nothing Then
        Set wsRealCD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsRealCD.Name = "ConsolidatedData"
    End If

    m_tempFolder = folder
    m_readSecs = -1
    ReadTempHeader CStr(cpPaths(1))
    ScanTextColumns CStr(cpPaths(1)), 5000
    ScanTextColumns CStr(ddPaths(1)), 5000

    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    FinishLegacyExport wsHome, wsRealCD, ecmID, AlertID, saveFolderPath, origCalc, _
        cpPaths, ddPaths, alPaths, "Continued from the export started " & stoppedAt & "."
    Exit Sub

ResumeFailed:
    Dim savedErrNum As Long, savedErrDesc As String
    savedErrNum = Err.Number
    savedErrDesc = Err.Description
    On Error GoTo -1
    On Error Resume Next
    CloseUnsavedExports
    Close
    CloseCsvHandle
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
        MsgBox "Process Safely Cancelled. The temporary files are kept - run " & _
            "Consolidated_AML_Workflow_Large_Resume again to finish.", vbInformation, "Aborted"
    ElseIf savedErrNum <> 0 And savedErrNum <> ERR_REPORTED Then
        MsgBox "An unexpected error occurred while " & m_stage & ":" & vbCrLf & vbCrLf & _
            "Error " & savedErrNum & ": " & savedErrDesc & vbCrLf & vbCrLf & _
            "The temporary files are kept - run Consolidated_AML_Workflow_Large_Resume again " & _
            "to finish.", vbCritical, "Large Export Error"
    End If
End Sub

' ==========================================================
' PASS 2 - lists, ConsolidatedData, Sheet7, main file
' ==========================================================
' Shared by LegacyExport and the Resume macro. Needs the column set-up
' (m_headerNames, m_outCols, m_dateCol, m_cpTextCol, m_textCol) and
' m_tempFolder; summary goes into the completion message.
' ==========================================================
Private Sub FinishLegacyExport(ByVal wsHome As Worksheet, ByVal wsRealCD As Worksheet, _
    ByVal ecmID As String, ByVal AlertID As String, ByVal saveFolderPath As String, _
    ByVal origCalc As XlCalculation, ByVal cpPaths As Collection, ByVal ddPaths As Collection, _
    ByVal alPaths As Collection, ByVal summary As String)

    Dim slash As String, prefix As String, stagePrefix As String
    Dim dateName As String, amtName As String, cdFilled As Boolean, bigCase As Boolean
    Dim cpStaged As Collection, ddStaged As Collection, cpRows As Long, ddRows As Long, alertRows As Long
    Dim cpFilesText As String, ddFilesText As String, unmatched As String, doneMsg As String
    Dim st(1 To 11) As Variant, newWb As Workbook, ws As Worksheet, wsScn As Worksheet, wsDay As Worksheet
    Dim wsPivot As Worksheet, outArr() As Variant, g As Long, shIdx As Long
    Dim excelFileName As String, finalSavePath As String
    Dim t0 As Double, tCp As Double, tDd As Double, tRest As Double, timing As String

    slash = Application.PathSeparator
    prefix = saveFolderPath & slash & ecmID & "_" & AlertID & "_"
    stagePrefix = m_tempFolder & slash & ecmID & "_" & AlertID & "_"

    If m_dateCol > 0 Then dateName = m_headerNames(m_dateCol)
    m_iAmt = FindHeaderIndex(m_headerNames, "Transaction Amount", False)
    m_iDr = FindHeaderIndex(m_headerNames, "Dr Cr", False)
    m_iAlert = FindHeaderIndex(m_headerNames, "Alert Information", False)
    m_iAcct = FindHeaderIndex(m_headerNames, "Account No", False)
    If m_iAmt > 0 Then amtName = m_headerNames(m_iAmt)
    ResetTotals

    ' ---- the lists, one .xlsx per temporary file, built in the scratch folder
    ' DeDupe on one file: ConsolidatedData = DeDupe, as in Module9 (copied
    ' while that file is open).
    Set cpStaged = New Collection
    Set ddStaged = New Collection
    wsRealCD.Cells.Clear
    t0 = Timer
    cpFilesText = SaveTempPartsAsXlsx(cpPaths, "CP Selection", stagePrefix & "CP Selection", dateName, _
        "dddd, mmmm d, yyyy", amtName, 1, Nothing, cdFilled, cpStaged, cpRows)
    tCp = Timer - t0
    t0 = Timer
    ddFilesText = SaveTempPartsAsXlsx(ddPaths, "DeDupe", stagePrefix & "DeDupe", dateName, _
        "m/d/yyyy", amtName, 2, wsRealCD, cdFilled, ddStaged, ddRows)
    tDd = Timer - t0
    t0 = Timer

    ' ---- ConsolidatedData + Sheet7
    ' Too big for one sheet: ConsolidatedData gets the rows that carry an
    ' alert (every rule name Generate Narrative looks up is on those), and
    ' Sheet7's figures from every DeDupe transaction.
    SetStage "updating ConsolidatedData"
    bigCase = Not cdFilled
    If bigCase Then
        CopyTempPartToSheet alPaths, wsRealCD
        alertRows = LastDataRow(wsRealCD) - 1
        If alertRows < 0 Then alertRows = 0
        st(1) = m_stCount: st(2) = m_stSum
        If m_stHaveDate Then
            st(3) = m_stFirst: st(4) = m_stLast
        Else
            st(3) = "01/00/1900": st(4) = "01/00/1900"   ' TEXT(0,"mm/dd/yyyy"), as MIN of nothing gives
        End If
        st(5) = m_acctText
        If m_stHavePos Then
            st(6) = m_stMinPos: st(7) = m_stMaxPos
        Else
            st(6) = 0: st(7) = 0
        End If
        st(8) = m_crSum: st(9) = m_drSum: st(10) = m_crCnt: st(11) = m_drCnt
        WriteLargeCaseStats st, wsHome.Range("J9").Value
        unmatched = PointSheet7AtLargeStats()
    Else
        ClearLargeCaseStats
    End If
    FormatDataColumns wsRealCD, dateName, "m/d/yyyy", amtName
    TidyDataSheet wsRealCD

    On Error Resume Next
    Module3.RefreshRuleNameTag
    On Error GoTo 0

    ' ---- main file: totals + pivots
    SetStage "writing the pivot totals"
    If m_scnN + 1 > wsRealCD.Rows.count Then
        Err.Raise vbObjectError + 1006, "modLargeExport", "There are too many alert / Dr Cr / counterparty " & _
            "combinations (" & Format$(m_scnN, "#,##0") & ") for the Scenario Totals sheet."
    End If
    Set newWb = NewOutputWorkbook()
    Set wsScn = newWb.Sheets(1)
    wsScn.Name = "Scenario Totals"
    wsScn.Columns(3).NumberFormat = "@"     ' Counterparty as text, as its formula gave it
    ReDim outArr(1 To m_scnN + 1, 1 To 5)
    outArr(1, 1) = "Alert Information": outArr(1, 2) = "Dr Cr": outArr(1, 3) = "Counterparty"
    outArr(1, 4) = "Transaction Amount": outArr(1, 5) = "Transaction Count"
    For g = 1 To m_scnN
        outArr(g + 1, 1) = m_scnAlert(g): outArr(g + 1, 2) = m_scnDr(g): outArr(g + 1, 3) = m_scnCp(g)
        outArr(g + 1, 4) = m_scnSum(g): outArr(g + 1, 5) = m_scnCnt(g)
    Next g
    wsScn.Range("A1").Resize(m_scnN + 1, 5).Value = outArr

    Set wsDay = newWb.Sheets.Add(After:=wsScn)
    wsDay.Name = "Daily Totals"
    ReDim outArr(1 To m_dayN + 1, 1 To 4)
    outArr(1, 1) = "Transaction Date": outArr(1, 2) = "Dr Cr"
    outArr(1, 3) = "Transaction Amount": outArr(1, 4) = "Transaction Count"
    For g = 1 To m_dayN
        outArr(g + 1, 1) = m_dayDate(g): outArr(g + 1, 2) = m_dayDr(g)
        outArr(g + 1, 3) = m_daySum(g): outArr(g + 1, 4) = m_dayCnt(g)
    Next g
    wsDay.Range("A1").Resize(m_dayN + 1, 4).Value = outArr

    ' Backwards, because the new workbook's other blank sheets go as it runs.
    For shIdx = newWb.Worksheets.count To 1 Step -1
        Set ws = newWb.Worksheets(shIdx)
        If ws.Name = "Scenario Totals" Or ws.Name = "Daily Totals" Then
            StyleHeaderRow ws
            FormatDataColumns ws, "Transaction Date", "m/d/yyyy", "Transaction Amount"
            TidyDataSheet ws
        Else
            SafeDeleteSheet newWb, ws.Name
        End If
    Next shIdx

    SetStage "building the Legacy pivots"
    Set wsPivot = newWb.Sheets.Add(Before:=newWb.Sheets(1))
    wsPivot.Name = "Pivot"
    BuildTotalsPivots newWb, wsPivot, wsScn, wsDay

    ' ---- everything is built: into the case folder together
    SetStage "moving the files into the case folder"
    MoveStagedFiles cpStaged, prefix & "CP Selection"
    MoveStagedFiles ddStaged, prefix & "DeDupe"
    m_keepTemp = False
    RemoveTempFolder

    excelFileName = ecmID & "_" & AlertID & "_Combined_Alerted_Transaction.xlsx"
    finalSavePath = saveFolderPath & slash & excelFileName
    SetStage "saving the Legacy file"
    CloseIfAlreadyOpen finalSavePath
    MoveExistingExportAside finalSavePath
    ZoomAllSheets newWb
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
    MarkSaved newWb
    DiscardPreviousExport finalSavePath

    newWb.Sheets("Pivot").Activate
    FinishRun origCalc

    tRest = Timer - t0
    If m_readSecs >= 0 Then timing = "reading " & DurationText(m_readSecs) & ", "
    timing = "Time taken: " & timing & "CP Selection " & DurationText(tCp) & ", DeDupe " & _
        DurationText(tDd) & ", pivots and the rest " & DurationText(tRest) & "."
    If m_readSecs >= 0 Then
        timing = timing & " Total " & DurationText(m_readSecs + IIf(tCp < 0, tCp + 86400, tCp) + _
            IIf(tDd < 0, tDd + 86400, tDd) + IIf(tRest < 0, tRest + 86400, tRest)) & "."
    End If

    doneMsg = "Workflow Complete!" & vbCrLf & vbCrLf & _
        "Pivots and totals:" & vbCrLf & "   " & excelFileName & vbCrLf & _
        "CP Selection - " & Format$(cpRows, "#,##0") & " rows:" & vbCrLf & "   " & cpFilesText & vbCrLf & _
        "DeDupe - " & Format$(ddRows, "#,##0") & " rows:" & vbCrLf & "   " & ddFilesText & vbCrLf & vbCrLf & _
        "All in: " & saveFolderPath & vbCrLf & _
        "(built first, then moved in together, so OneDrive syncs them once they're all there)" & _
        vbCrLf & vbCrLf & summary & vbCrLf & vbCrLf & timing
    If bigCase Then
        doneMsg = doneMsg & vbCrLf & vbCrLf & _
            "DeDupe is too big for ConsolidatedData, so it holds only the " & Format$(alertRows, "#,##0") & _
            " rows that carry an alert. Sheet7's narrative figures (count, totals, date range, CR/DR, " & _
            "account numbers) were worked out from all " & Format$(ddRows, "#,##0") & _
            " transactions, and apply while Sheet1 has this ECM ID."
        If Len(unmatched) > 0 Then
            doneMsg = doneMsg & vbCrLf & vbCrLf & "Check Sheet7 " & unmatched & ": these read " & _
                "ConsolidatedData but weren't recognised, so they only see the alerted rows."
        End If
    End If
    MsgBox doneMsg, vbInformation, "Success"
End Sub

' ==========================================================
' Pass 2's totals and Sheet7 figures, added up from each list file
' ==========================================================
' Scenario totals (per Alert Information / Dr Cr / Counterparty) come from
' the CP Selection files; Daily totals (per day / Dr Cr) and Sheet7's
' figures from the DeDupe files - the sheets Module9's pivots and
' ConsolidatedData read. Each file is added in while it is open for saving.
' ==========================================================
Private Sub ResetTotals()
    m_scnIdx = NewShards(): m_dayIdx = NewShards(): m_acctSeen = NewShards()
    m_scnN = 0: m_dayN = 0
    m_scnCap = 4096: m_dayCap = 4096
    ReDim m_scnAlert(1 To m_scnCap): ReDim m_scnDr(1 To m_scnCap): ReDim m_scnCp(1 To m_scnCap)
    ReDim m_scnSum(1 To m_scnCap): ReDim m_scnCnt(1 To m_scnCap)
    ReDim m_dayDate(1 To m_dayCap): ReDim m_dayDr(1 To m_dayCap)
    ReDim m_daySum(1 To m_dayCap): ReDim m_dayCnt(1 To m_dayCap)
    m_stCount = 0: m_stSum = 0: m_stHaveDate = False: m_stHavePos = False
    m_stMinPos = 0: m_stMaxPos = 0
    m_crSum = 0: m_drSum = 0: m_crCnt = 0: m_drCnt = 0
    m_acctText = ""
End Sub

' Column c of a list sheet, rows 2 to lastR, as a 2-D array (Empty when c = 0).
Private Function ColumnValues(ByVal ws As Worksheet, ByVal c As Long, ByVal lastR As Long) As Variant
    Dim v As Variant
    If c = 0 Or lastR < 2 Then Exit Function
    v = ws.Range(ws.Cells(2, c), ws.Cells(lastR, c)).Value
    If Not IsArray(v) Then v = OneCellArray(v)
    ColumnValues = v
End Function

' Row r of a ColumnValues array. ByRef so the array isn't copied on every
' call - the mistake that made Module9's ArrCell never finish at 6 lakh rows.
Private Function ColCell(ByRef a As Variant, ByVal r As Long) As Variant
    If IsArray(a) Then ColCell = a(r, 1)
End Function

Private Sub AddScenarioTotals(ByVal ws As Worksheet, ByVal nRows As Long)
    Dim aAlert As Variant, aDr As Variant, aCp As Variant, aAmt As Variant
    Dim r As Long, g As Long, prevN As Long, k As String, amt As Variant
    If nRows < 1 Then Exit Sub
    aAlert = ColumnValues(ws, m_iAlert, nRows + 1)
    aDr = ColumnValues(ws, m_iDr, nRows + 1)
    aCp = ColumnValues(ws, m_cpTextCol, nRows + 1)
    aAmt = ColumnValues(ws, m_iAmt, nRows + 1)
    For r = 1 To nRows
        k = KeyText(ColCell(aAlert, r)) & vbTab & KeyText(ColCell(aDr, r)) & vbTab & KeyText(ColCell(aCp, r))
        prevN = m_scnN
        g = GroupIndex(m_scnIdx, k, m_scnN)
        If m_scnN > prevN Then
            If m_scnN > m_scnCap Then
                m_scnCap = m_scnCap * 2
                ReDim Preserve m_scnAlert(1 To m_scnCap): ReDim Preserve m_scnDr(1 To m_scnCap)
                ReDim Preserve m_scnCp(1 To m_scnCap): ReDim Preserve m_scnSum(1 To m_scnCap)
                ReDim Preserve m_scnCnt(1 To m_scnCap)
            End If
            m_scnAlert(g) = ColCell(aAlert, r)
            m_scnDr(g) = ColCell(aDr, r)
            m_scnCp(g) = ColCell(aCp, r)
        End If
        amt = ColCell(aAmt, r)
        If IsNumber(amt) Then m_scnSum(g) = m_scnSum(g) + amt
        If Not IsBlankValue(amt) Then m_scnCnt(g) = m_scnCnt(g) + 1
    Next r
End Sub

Private Sub AddDailyTotalsAndStats(ByVal ws As Worksheet, ByVal nRows As Long)
    Dim aDate As Variant, aDr As Variant, aAmt As Variant, aAcct As Variant
    Dim r As Long, g As Long, prevN As Long, k As String, amt As Variant, dv As Variant
    Dim side As Variant, v As Variant
    If nRows < 1 Then Exit Sub
    aDate = ColumnValues(ws, m_dateCol, nRows + 1)
    aDr = ColumnValues(ws, m_iDr, nRows + 1)
    aAmt = ColumnValues(ws, m_iAmt, nRows + 1)
    aAcct = ColumnValues(ws, m_iAcct, nRows + 1)
    For r = 1 To nRows
        amt = ColCell(aAmt, r)
        side = ColCell(aDr, r)

        ' Daily totals: per day / Dr Cr
        dv = ColCell(aDate, r)
        If VarType(dv) = vbDate Then
            dv = CDate(Int(CDbl(dv)))
            k = "D" & CStr(CLng(CDbl(dv)))
        Else
            k = "T" & KeyText(dv)
        End If
        k = KeyText(side) & vbTab & k
        prevN = m_dayN
        g = GroupIndex(m_dayIdx, k, m_dayN)
        If m_dayN > prevN Then
            If m_dayN > m_dayCap Then
                m_dayCap = m_dayCap * 2
                ReDim Preserve m_dayDate(1 To m_dayCap): ReDim Preserve m_dayDr(1 To m_dayCap)
                ReDim Preserve m_daySum(1 To m_dayCap): ReDim Preserve m_dayCnt(1 To m_dayCap)
            End If
            m_dayDate(g) = dv
            m_dayDr(g) = side
        End If
        If IsNumber(amt) Then m_daySum(g) = m_daySum(g) + amt
        If Not IsBlankValue(amt) Then m_dayCnt(g) = m_dayCnt(g) + 1

        ' Sheet7's figures (what its formulas give on ConsolidatedData = DeDupe)
        If IsNumber(amt) Then
            m_stCount = m_stCount + 1
            m_stSum = m_stSum + amt
            If amt > 0 Then
                If Not m_stHavePos Then
                    m_stMinPos = amt: m_stMaxPos = amt: m_stHavePos = True
                Else
                    If amt < m_stMinPos Then m_stMinPos = amt
                    If amt > m_stMaxPos Then m_stMaxPos = amt
                End If
            End If
        End If
        If VarType(dv) = vbDate Then
            If Not m_stHaveDate Then
                m_stFirst = dv: m_stLast = dv: m_stHaveDate = True
            Else
                If dv < m_stFirst Then m_stFirst = dv
                If dv > m_stLast Then m_stLast = dv
            End If
        End If
        If VarType(side) = vbString Then
            If StrComp(side, "CR", vbTextCompare) = 0 Then
                m_crCnt = m_crCnt + 1
                If IsNumber(amt) Then m_crSum = m_crSum + amt
            ElseIf StrComp(side, "DR", vbTextCompare) = 0 Then
                m_drCnt = m_drCnt + 1
                If IsNumber(amt) Then m_drSum = m_drSum + amt
            End If
        End If
        If Len(m_acctText) < 32000 Then
            v = ColCell(aAcct, r)
            If Not IsBlankValue(v) And Not IsError(v) Then
                If Not AlreadySeen(m_acctSeen, KeyText(v)) Then
                    If Len(m_acctText) > 0 Then m_acctText = m_acctText & ","
                    m_acctText = m_acctText & CStr(v)
                End If
            End If
        End If
    Next r
End Sub

' ==========================================================
' Temporary files - pass 1's kept rows, on disk until pass 2
' ==========================================================
' Each TempWriter writes tab-separated UTF-8 text to "<BaseName> n.txt" in
' the scratch folder: the header line, then up to RAW_ROWS_PER_SHEET rows
' per file, carrying on into the next file when one fills. Lines collect in
' Lines and are written WRITE_BLOCK_ROWS at a time.
' ==========================================================
Private Sub TempOpen(ByRef w As TempWriter, ByVal baseName As String)
    w.BaseName = baseName
    w.Fill = 0
    w.Total = 0
    w.FileNo = 0
    Set w.Paths = New Collection
    ReDim w.Lines(1 To WRITE_BLOCK_ROWS)
End Sub

Private Sub TempStartFile(ByRef w As TempWriter)
    Dim p As String, fn As Integer
    p = m_tempFolder & Application.PathSeparator & w.BaseName & " " & (w.Paths.count + 1) & ".txt"
    fn = FreeFile
    Open p For Binary Access Write As #fn
    w.FileNo = fn
    w.Paths.Add p
    w.PartRows = 0
    PutUtf8 fn, m_headerLine & vbCrLf
End Sub

Private Sub TempAdd(ByRef w As TempWriter, ByRef rowText As String)
    w.Fill = w.Fill + 1
    w.Lines(w.Fill) = rowText
    If w.Fill = WRITE_BLOCK_ROWS Then TempFlush w
End Sub

Private Sub TempFlush(ByRef w As TempWriter)
    Dim done As Long, n As Long, room As Long, chunk() As String, i As Long, fn As Integer
    Do While done < w.Fill
        If w.FileNo = 0 Then TempStartFile w
        room = RAW_ROWS_PER_SHEET - w.PartRows
        If room <= 0 Then
            fn = w.FileNo
            Close #fn
            TempStartFile w
            room = RAW_ROWS_PER_SHEET
        End If
        n = w.Fill - done
        If n > room Then n = room
        If done = 0 And n = WRITE_BLOCK_ROWS Then
            PutUtf8 w.FileNo, Join(w.Lines, vbCrLf) & vbCrLf
        Else
            ReDim chunk(1 To n)
            For i = 1 To n
                chunk(i) = w.Lines(done + i)
            Next i
            PutUtf8 w.FileNo, Join(chunk, vbCrLf) & vbCrLf
        End If
        w.PartRows = w.PartRows + n
        w.Total = w.Total + n
        done = done + n
    Loop
    w.Fill = 0
End Sub

Private Sub TempClose(ByRef w As TempWriter)
    Dim fn As Integer
    TempFlush w
    If w.FileNo <> 0 Then
        fn = w.FileNo
        Close #fn
        w.FileNo = 0
    End If
End Sub

' Writes text to an open binary file as UTF-8 (so accented names survive),
' without the byte-order mark ADODB puts at the start.
Private Sub PutUtf8(ByVal fn As Integer, ByVal s As String)
    Dim stm As Object, b() As Byte
    If Len(s) = 0 Then Exit Sub
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                ' text
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.Position = 0
    stm.Type = 1                ' binary
    stm.Position = 3            ' past the BOM
    b = stm.Read
    stm.Close
    Put #fn, , b
End Sub

' One export row as a line of a temporary file.
Private Function RowLine(ByRef rowVals() As Variant) As String
    Dim j As Long
    For j = 1 To m_outCols
        m_fields(j) = FieldText(rowVals(j), j)
    Next j
    RowLine = Join(m_fields, vbTab)
End Function

' A value written so Excel's text import reads it back as the same value:
' dates as mm/dd/yyyy (the date column is imported month/day/year),
' numbers with a "." decimal and no separators, text with any tab or line
' break turned into a space (they would split the row). A column holding
' digit strings longer than 15 is marked to import as text, because as a
' number Excel keeps only 15 digits and the rest would be lost.
Private Function FieldText(ByRef v As Variant, ByVal j As Long) As String
    Select Case VarType(v)
        Case vbEmpty, vbNull, vbError
            ' left empty
        Case vbString
            FieldText = CleanFieldText(CStr(v))
            If Len(FieldText) > 15 Then
                If Not m_textCol(j) Then
                    If Not (FieldText Like "*[!0-9]*") Then m_textCol(j) = True
                End If
            End If
        Case vbDate
            If v = Int(v) Then
                FieldText = Format$(v, "mm\/dd\/yyyy")
            Else
                FieldText = Format$(v, "mm\/dd\/yyyy hh:nn:ss")
            End If
        Case vbBoolean
            If v Then FieldText = "TRUE" Else FieldText = "FALSE"
        Case Else
            FieldText = Trim$(Str$(v))
    End Select
End Function

Private Function CleanFieldText(ByVal s As String) As String
    If InStr(s, vbTab) > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        s = Replace(Replace(Replace(Replace(s, vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
    End If
    CleanFieldText = s
End Function

' How Excel reads each column back: the date column month/day/year,
' Counterparty and over-long digit columns as text, the rest as Excel
' would read them typed in (as Module9's value rewrite left them).
Private Function ImportFieldInfo() As Variant
    Dim fi() As Variant, j As Long, t As Long
    ReDim fi(0 To m_outCols - 1)
    For j = 1 To m_outCols
        t = xlGeneralFormat
        If j = m_dateCol Then t = xlMDYFormat
        If j = m_cpTextCol Or m_textCol(j) Then t = xlTextFormat
        fi(j - 1) = Array(j, t)
    Next j
    ImportFieldInfo = fi
End Function

' Opens a temporary file in Excel as a workbook of its own.
Private Function OpenTempPart(ByVal p As String) As Workbook
    Workbooks.OpenText fileName:=p, Origin:=65001, StartRow:=1, DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierNone, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
        FieldInfo:=ImportFieldInfo(), Local:=False
    Set OpenTempPart = ActiveWorkbook
    m_unsaved.Add OpenTempPart
End Function

' Pass 2 for one list: each temporary file becomes basePath.xlsx, or
' basePath (part 1).xlsx, (part 2)... when there is more than one, each
' saved and closed before the next is opened. basePath is in the scratch
' folder; the files made are added to staged for MoveStagedFiles, and the
' temporary files are left for RemoveTempFolder (so a failure here can
' still be resumed). totalsMode 1 adds each file into the Scenario totals,
' 2 into the Daily totals and Sheet7's figures. copyTo: when the whole
' list fits in one file, it is also copied onto this sheet
' (ConsolidatedData) and copied comes back True. totalRows gets the list's
' row count. Returns the file names, for the completion message.
Private Function SaveTempPartsAsXlsx(ByVal paths As Collection, ByVal sheetName As String, _
    ByVal basePath As String, ByVal dateName As String, ByVal dateFormat As String, _
    ByVal amtName As String, ByVal totalsMode As Long, ByVal copyTo As Worksheet, _
    ByRef copied As Boolean, ByVal staged As Collection, ByRef totalRows As Long) As String
    Dim p As Long, n As Long, nRows As Long, wb As Workbook, ws As Worksheet
    Dim target As String, fileList As String

    n = paths.count
    For p = 1 To n
        SetStage "saving " & sheetName & IIf(n > 1, " (part " & p & " of " & n & ")", "")
        DoEvents
        Set wb = OpenTempPart(CStr(paths(p)))
        Set ws = wb.Worksheets(1)
        ws.Name = sheetName

        nRows = LastDataRow(ws) - 1
        If nRows < 0 Then nRows = 0
        totalRows = totalRows + nRows
        If totalsMode = 1 Then
            AddScenarioTotals ws, nRows
        Else
            AddDailyTotalsAndStats ws, nRows
        End If

        StyleHeaderRow ws
        FormatDataColumns ws, dateName, dateFormat, amtName
        TidyDataSheet ws

        If n = 1 And Not copyTo Is Nothing Then
            ws.UsedRange.Copy Destination:=copyTo.Range("A1")
            copied = True
        End If

        If n = 1 Then target = basePath & LIST_FILE_EXT Else target = basePath & " (part " & p & ")" & LIST_FILE_EXT
        SetSheetZoom85 wb, Array(ws.Name)
        Application.DisplayAlerts = False
        wb.SaveAs fileName:=target, FileFormat:=LIST_FILE_FORMAT
        MarkSaved wb
        wb.Close SaveChanges:=False
        staged.Add target

        If Len(fileList) > 0 Then fileList = fileList & vbCrLf & "   "
        fileList = fileList & Mid$(target, InStrRev(target, Application.PathSeparator) + 1)
    Next p
    If n = 0 Then fileList = "(none)"
    SaveTempPartsAsXlsx = fileList
End Function

' The first temporary file of a list, onto a sheet (the alerted rows onto
' ConsolidatedData).
Private Sub CopyTempPartToSheet(ByVal paths As Collection, ByVal target As Worksheet)
    Dim wb As Workbook
    If paths.count = 0 Then Exit Sub
    Set wb = OpenTempPart(CStr(paths(1)))
    wb.Worksheets(1).UsedRange.Copy Destination:=target.Range("A1")
    MarkSaved wb
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
End Sub

' Moves a list's finished files from the scratch folder into the case
' folder as finalBase.xlsx / finalBase (part n).xlsx, after removing the
' files an earlier run left there under finalBase - so a smaller re-run
' leaves no stale parts behind.
Private Sub MoveStagedFiles(ByVal staged As Collection, ByVal finalBase As String)
    Dim FSO As Object, item As Variant, folder As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    folder = Left$(finalBase, InStrRev(finalBase, Application.PathSeparator))
    RemoveOldExportFiles finalBase
    For Each item In staged
        FSO.MoveFile CStr(item), folder & FileNameOf(CStr(item))
    Next item
End Sub

' Deletes basePath and basePath (part n) files - .xlsx or .xlsb - from an
' earlier run, closing any of them that are open in Excel.
Private Sub RemoveOldExportFiles(ByVal basePath As String)
    Dim folder As String, f As String, found As Collection, item As Variant, ext As Variant
    folder = Left$(basePath, InStrRev(basePath, Application.PathSeparator))
    Set found = New Collection
    For Each ext In Array(".xlsx", ".xlsb")
        If Len(Dir(basePath & ext)) > 0 Then found.Add basePath & ext
        f = Dir(basePath & " (part *)" & ext)
        Do While Len(f) > 0
            found.Add folder & f
            f = Dir()
        Loop
    Next ext
    For Each item In found
        CloseIfAlreadyOpen CStr(item)
        Kill CStr(item)
    Next item
End Sub

' Removes the scratch folder and everything in it.
Private Sub RemoveTempFolder()
    On Error Resume Next
    If Len(m_tempFolder) = 0 Then Exit Sub
    Kill m_tempFolder & Application.PathSeparator & "*.*"
    RmDir m_tempFolder
    m_tempFolder = ""
    On Error GoTo 0
End Sub

' The newest scratch folder whose pass 1 finished (it has a DeDupe file),
' or "" if there is none.
Private Function LatestUnfinishedExport() As String
    Dim root As String, d As String, best As String
    root = Environ$("TEMP") & Application.PathSeparator
    d = Dir(root & "LargeExport_*", vbDirectory)
    Do While Len(d) > 0
        If (GetAttr(root & d) And vbDirectory) = vbDirectory Then
            If Len(Dir(root & d & Application.PathSeparator & "DeDupe 1.txt")) > 0 Then
                If d > best Then best = d
            End If
        End If
        d = Dir()
    Loop
    If Len(best) > 0 Then LatestUnfinishedExport = root & best
End Function

' "<baseName> 1.txt", "<baseName> 2.txt"... in a scratch folder, in order.
Private Function TempPartPaths(ByVal folder As String, ByVal baseName As String) As Collection
    Dim c As New Collection, n As Long, p As String
    Do
        n = n + 1
        p = folder & Application.PathSeparator & baseName & " " & n & ".txt"
        If Len(Dir(p)) = 0 Then Exit Do
        c.Add p
    Loop
    Set TempPartPaths = c
End Function

' Sets the column set-up from a temporary file's header line (the Resume
' macro's stand-in for pass 1's).
Private Sub ReadTempHeader(ByVal p As String)
    Dim fn As Integer, s As String, parts() As String, j As Long
    fn = FreeFile
    Open p For Input As #fn
    Line Input #fn, s
    Close #fn
    parts = Split(s, vbTab)
    m_outCols = UBound(parts) + 1
    ReDim m_headerNames(1 To m_outCols)
    ReDim m_fields(1 To m_outCols)
    ReDim m_textCol(1 To m_outCols)
    For j = 1 To m_outCols
        m_headerNames(j) = parts(j - 1)
    Next j
    m_headerLine = s
    m_dateCol = FindHeaderIndex(m_headerNames, "Transaction Date", False)
    m_cpTextCol = 0
    If m_headerNames(m_outCols) = "Counterparty" Then m_cpTextCol = m_outCols
End Sub

' Marks the columns to import as text from the first maxLines rows of a
' temporary file - pass 1 marks them as it writes; the Resume macro
' re-checks them here.
Private Sub ScanTextColumns(ByVal p As String, ByVal maxLines As Long)
    Dim fn As Integer, s As String, parts() As String, n As Long, j As Long
    fn = FreeFile
    Open p For Input As #fn
    If Not EOF(fn) Then Line Input #fn, s          ' header
    Do While Not EOF(fn) And n < maxLines
        Line Input #fn, s
        n = n + 1
        parts = Split(s, vbTab)
        For j = 0 To UBound(parts)
            If j < m_outCols Then
                If Len(parts(j)) > 15 Then
                    If Not (parts(j) Like "*[!0-9]*") Then m_textCol(j + 1) = True
                End If
            End If
        Next j
    Loop
    Close #fn
End Sub

' Bold white on blue, the header look the exports carried from the source.
Private Sub StyleHeaderRow(ByVal ws As Worksheet)
    Dim lastC As Long
    lastC = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastC))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
    End With
End Sub

' Duplicate checks and totals use Scripting.Dictionary, split into SHARDS
' dictionaries by a hash of the key so no one dictionary holds lakhs of
' keys - a single one slows down badly well before 30 lakh.
Private Function NewShards() As Variant
    Dim a() As Object, i As Long
    ReDim a(0 To SHARDS - 1)
    For i = 0 To SHARDS - 1
        Set a(i) = CreateObject("Scripting.Dictionary")
    Next i
    NewShards = a
End Function

Private Function ShardOf(ByRef k As String) As Long
    Dim n As Long, h As Long, i As Long, stopAt As Long
    n = Len(k)
    h = n
    stopAt = n - 5
    If stopAt < 1 Then stopAt = 1
    For i = n To stopAt Step -1
        h = (h * 31 + (AscW(Mid$(k, i, 1)) And &HFFFF&)) Mod 1000003
    Next i
    ShardOf = h Mod SHARDS
End Function

' True if k was seen before; otherwise remembers it and returns False.
Private Function AlreadySeen(ByRef shards As Variant, ByRef k As String) As Boolean
    Dim d As Object
    Set d = shards(ShardOf(k))
    If d.Exists(k) Then
        AlreadySeen = True
    Else
        d.Add k, Empty
    End If
End Function

' The number of k's group, numbering a new group n + 1 (and counting it).
Private Function GroupIndex(ByRef shards As Variant, ByVal k As String, ByRef n As Long) As Long
    Dim d As Object
    Set d = shards(ShardOf(k))
    If d.Exists(k) Then
        GroupIndex = d(k)
    Else
        n = n + 1
        d.Add k, n
        GroupIndex = n
    End If
End Function

' Duplicate-check text for a value: upper case (RemoveDuplicates ignores
' case) and the number 1877576904 the same as the text "1877576904" (as
' they were once Module9's value rewrite made both numbers).
Private Function KeyText(ByRef v As Variant) As String
    If IsError(v) Then
        KeyText = "#ERROR"
    ElseIf Not IsEmpty(v) Then
        KeyText = UCase$(CStr(v))
    End If
End Function

Private Function IsBlankValue(ByRef v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Len(v) = 0)
    End If
End Function

Private Function IsNumber(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbInteger, vbLong, vbByte
            IsNumber = True
    End Select
End Function

' Transaction Date as Module9's TextToColumns (month/day/year) leaves it:
' dates stay dates, date serial numbers become dates, text dates are read
' by ParseMdyText, anything else is left as it was.
Private Function CleanDate(ByVal v As Variant) As Variant
    CleanDate = v
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbInteger, vbLong
            If v >= 1 And v < 2958466 Then CleanDate = CDate(v)
        Case vbString
            CleanDate = ParseMdyText(CStr(v))
    End Select
End Function

' m/d/yyyy or m-d-yyyy (2-digit years: 00-29 = 2000s, 30-99 = 1900s, as
' Excel reads them) or yyyy-mm-dd, each with an optional time after a
' space; dates written with a month name go through VBA's own date
' reading. Blank text becomes empty; anything else stays as text, as
' TextToColumns leaves a date it can't read.
Private Function ParseMdyText(ByVal s As String) As Variant
    Dim t As String, datePart As String, timePart As String, p() As String, sp As Long
    Dim y As Long, mo As Long, d As Long, dt As Date

    ParseMdyText = s
    t = Trim$(s)
    If Len(t) = 0 Then
        ParseMdyText = Empty
        Exit Function
    End If
    sp = InStr(t, " ")
    If sp > 0 Then
        datePart = Left$(t, sp - 1)
        timePart = Trim$(Mid$(t, sp + 1))
    Else
        datePart = t
    End If

    If datePart Like "*[A-Za-z]*" Then
        If IsDate(t) Then ParseMdyText = CDate(t)
        Exit Function
    End If

    p = Split(Replace(datePart, "-", "/"), "/")
    If UBound(p) <> 2 Then Exit Function
    If Not (IsShortDigits(p(0)) And IsShortDigits(p(1)) And IsShortDigits(p(2))) Then Exit Function
    If Len(p(0)) = 4 Then
        y = CLng(p(0)): mo = CLng(p(1)): d = CLng(p(2))
    Else
        mo = CLng(p(0)): d = CLng(p(1)): y = CLng(p(2))
        If Len(p(2)) <= 2 Then
            If y < 30 Then y = y + 2000 Else y = y + 1900
        End If
    End If
    If mo < 1 Or mo > 12 Or d < 1 Or d > 31 Or y < 1900 Or y > 9999 Then Exit Function
    dt = DateSerial(y, mo, d)
    If Month(dt) <> mo Or Day(dt) <> d Then Exit Function
    If Len(timePart) > 0 Then
        If Not IsDate(timePart) Then Exit Function
        dt = dt + TimeValue(timePart)
    End If
    ParseMdyText = dt
End Function

Private Function IsShortDigits(ByVal s As String) As Boolean
    If Len(s) = 0 Or Len(s) > 4 Then Exit Function
    IsShortDigits = Not (s Like "*[!0-9]*")
End Function

' Transaction Amount as Module9's value rewrite leaves it: text holding a
' number ("$1,018.65", "(250.00)") becomes that number; blank text becomes
' empty; anything else is left as it was.
Private Function CleanAmount(ByVal v As Variant) As Variant
    Dim s As String, neg As Boolean
    CleanAmount = v
    If VarType(v) <> vbString Then Exit Function
    s = Trim$(v)
    If Len(s) = 0 Then
        CleanAmount = Empty
        Exit Function
    End If
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        neg = True
        s = Mid$(s, 2, Len(s) - 2)
    End If
    s = Replace(Replace(Replace(s, "$", ""), ",", ""), " ", "")
    If Not LooksNumeric(s) Then Exit Function
    If neg Then CleanAmount = -Val(s) Else CleanAmount = Val(s)
End Function

Private Function LooksNumeric(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9.-]*" Then Exit Function
    If s Like "*.*.*" Then Exit Function
    If InStr(2, s, "-") > 0 Then Exit Function
    LooksNumeric = (s Like "*#*")
End Function

' Row r, columns c1..c2, as text (errors and blanks as "").
Private Function RowTexts(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As String()
    Dim v As Variant, out() As String, i As Long, n As Long
    n = c2 - c1 + 1
    ReDim out(1 To n)
    v = ws.Range(ws.Cells(r, c1), ws.Cells(r, c2)).Value
    If n = 1 Then
        If Not IsError(v) Then out(1) = CStr(v)
    Else
        For i = 1 To n
            If Not IsError(v(1, i)) Then out(i) = CStr(v(1, i))
        Next i
    End If
    RowTexts = out
End Function

Private Function OneCellArray(ByVal v As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    a(1, 1) = v
    OneCellArray = a
End Function

' For each column of the export (master), which column of this file holds
' it: same name, any case, spaces around it ignored. The first file - and any file with the same
' header - maps straight across. Columns this file has that the export
' doesn't are counted into extraCols.
Private Function MapColumns(ByRef master() As String, ByRef fileHdr() As String, ByRef extraCols As Long) As Long()
    Dim m() As Long, used() As Boolean, i As Long, j As Long, same As Boolean
    ReDim m(1 To UBound(master))
    ReDim used(1 To UBound(fileHdr))

    same = (UBound(fileHdr) = UBound(master))
    If same Then
        For i = 1 To UBound(master)
            If StrComp(Trim$(master(i)), Trim$(fileHdr(i)), vbTextCompare) <> 0 Then
                same = False
                Exit For
            End If
        Next i
    End If

    If same Then
        For i = 1 To UBound(master)
            m(i) = i
        Next i
    Else
        For i = 1 To UBound(master)
            For j = 1 To UBound(fileHdr)
                If Not used(j) Then
                    If StrComp(Trim$(master(i)), Trim$(fileHdr(j)), vbTextCompare) = 0 Then
                        m(i) = j
                        used(j) = True
                        Exit For
                    End If
                End If
            Next j
        Next i
        For j = 1 To UBound(fileHdr)
            If Not used(j) And Len(fileHdr(j)) > 0 Then extraCols = extraCols + 1
        Next j
    End If
    MapColumns = m
End Function

' FindHeaderName's search, returning the column's position (0 if none).
Private Function FindHeaderIndex(ByRef names() As String, ByVal what As String, ByVal wholeMatch As Boolean) As Long
    Dim n As Long, k As Long, i As Long
    n = UBound(names) - LBound(names) + 1
    For k = 1 To n
        i = LBound(names) + (k Mod n)     ' 2nd, 3rd, ... last, then 1st - Rows(1).Find's order
        If wholeMatch Then
            If StrComp(names(i), what, vbTextCompare) = 0 Then FindHeaderIndex = i: Exit Function
        Else
            If InStr(1, names(i), what, vbTextCompare) > 0 Then FindHeaderIndex = i: Exit Function
        End If
    Next k
End Function

Private Function FileNameOf(ByVal path As String) As String
    FileNameOf = Mid$(path, InStrRev(path, Application.PathSeparator) + 1)
End Function
