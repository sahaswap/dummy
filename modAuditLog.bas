Attribute VB_Name = "modAuditLog"
'==================================================================
' modAuditLog  -  centralized cross-tool audit ledger (v3.6)
'
' Every macro in this workbook that produces a case-relevant
' artifact - folder creation (Module8), transaction consolidation
' (Module9), narrative/RFI generation (Module3 / modRFIGenerate),
' OSDD search (Module2) - calls LogAuditEvent once, right after it
' succeeds. That keeps a single "Audit_Ledger" sheet in THIS
' workbook as one chronological index of everything that happened
' to every case: filter it by ECM Case ID and you see that case's
' whole lifecycle in order, across every macro that touched it.
'
' Deliberately a DIFFERENT sheet name from the pre-existing
' "Audit_Log" tab some workbooks already have (the original v3.4
' OSINT-only ledger) - this module never touches, renames, or reads
' that sheet. It creates and owns "Audit_Ledger" only, so there's no
' chance of interacting with whatever's already in Audit_Log.
'
' LogAuditEvent also appends the same row to a "Register" tab inside
' a per-case audit workbook - Desktop\{ECMID}\{ECMID}_Audit_Log.xlsx -
' filtered to ONLY that case's rows, never the whole firm-wide log,
' so a folder that later gets shared onward (e.g. an RFI response to
' another bank) never carries any other customer's data with it.
'
' Two more entry points build out that same per-case workbook with
' real content, not just references:
'   ArchiveOutputSheets   - copies Module9's actual Raw Transactions/
'                           Pivot/CP Selection/DeDupe sheets in,
'                           tagged by source (Alerted/NonAlerted) so
'                           re-running the SAME tag refreshes those
'                           4 tabs while a DIFFERENT tag adds its own
'                           4 alongside rather than overwriting them.
'   ArchiveNarrativeText  - writes the full generated narrative/RFI
'                           text into a single "Narrative" tab,
'                           always the latest version (the Register
'                           tab's row history is what shows every
'                           time it was generated/regenerated).
'==================================================================
Option Explicit

Private Const AUDIT_SHEET_NAME As String = "Audit_Ledger"
Private Const AUDIT_XLSX_SUFFIX As String = "_Audit_Log.xlsx"
Private Const REGISTER_SHEET_NAME As String = "Register"
Private Const NARRATIVE_SHEET_NAME As String = "Narrative"

' Same protection password already used elsewhere in this workbook
' (Module9's sheet/workbook protect calls). Every sheet in the
' per-case audit workbook is protected, and the workbook structure
' is locked, before each save - so the file lands read-only to
' anyone opening it without the password, and only this module's own
' code (which knows it) can update it on the next event.
Private Const AUDIT_PROTECT_PASSWORD As String = "p7ss"

Private Const COL_TIMESTAMP As Long = 1
Private Const COL_ANALYST As Long = 2
Private Const COL_ECMCASE As Long = 3
Private Const COL_ALERTID As Long = 4
Private Const COL_CUSTOMER As Long = 5
Private Const COL_COUNTERPARTIES As Long = 6
Private Const COL_EVENTTYPE As Long = 7
Private Const COL_OUTPUTFILE As Long = 8
Private Const COL_TOOLVERSION As Long = 9
Private Const COL_NOTES As Long = 10
Private Const COL_DETAIL As Long = 11
Private Const AUDIT_HEADER_COLS As Long = 11

' Excel caps a cell at 32,767 characters. Detail is capped
' defensively so a freak oversized value can't throw a runtime
' error on the .Value write instead of just logging a truncated
' copy. (The full narrative text itself lives in the Narrative tab,
' one row per paragraph, which isn't subject to this cap the same
' way - see ArchiveNarrativeText.)
Private Const DETAIL_CELL_CAP As Long = 32000

'==================================================================
' PUBLIC ENTRY POINTS
'==================================================================

'------------------------------------------------------------------
' Call this right after any macro successfully produces a
' case-relevant artifact.
'
'   ecmID / alertID / customerName / counterparties - identify the
'       case. Pass whatever the calling macro already has in hand
'       (usually straight off Sheet1) rather than re-reading Sheet1
'       here, so this stays decoupled from Sheet1's exact layout.
'   eventType   - short label, e.g. "OSDD Search", "Transaction File
'                 Consolidated", "Narrative Generated".
'   outputFile  - full path of whatever file this event produced,
'                 or "" if the event didn't produce exactly one file.
'   toolVersion - the calling module's own TOOL_VERSION constant.
'   notes       - free-text detail (counts, error text, template
'                 name, etc.). Never parsed back out - purely for a
'                 human reading the log.
'   detail      - short-to-medium event-specific detail: the list of
'                 searches + mode used for an OSDD event, or a
'                 pointer to the Narrative tab for a narrative event
'                 (the full text itself goes through
'                 ArchiveNarrativeText, not through here). Blank for
'                 event types with no equivalent.
'   saveWorkbook - defaults True. Pass False from an error handler
'                 or anywhere ThisWorkbook.Save could hang (e.g. a
'                 blocked OneDrive sync) - matches the OSINT error
'                 path's existing "don't save on the error branch"
'                 behavior.
'------------------------------------------------------------------
Public Sub LogAuditEvent(ByVal ecmID As String, _
ByVal alertID As String, _
ByVal customerName As String, _
ByVal counterparties As String, _
ByVal eventType As String, _
ByVal outputFile As String, _
ByVal toolVersion As String, _
Optional ByVal notes As String = "", _
Optional ByVal detail As String = "", _
Optional ByVal saveWorkbook As Boolean = True)
On Error GoTo Fail

If Len(detail) > DETAIL_CELL_CAP Then
detail = Left$(detail, DETAIL_CELL_CAP) & " [TRUNCATED]"
End If

' 1. Firm-wide index, inside THIS workbook.
Dim ws As Worksheet
Set ws = GetOrCreateAuditSheet(ThisWorkbook)
WriteAuditRow ws, ecmID, alertID, customerName, counterparties, _
eventType, outputFile, toolVersion, notes, detail

If saveWorkbook Then
On Error Resume Next
ThisWorkbook.Save
On Error GoTo Fail
End If

' 2. This case's own Register tab, inside its Desktop\{ECMID}\
' audit workbook.
WriteCaseRegisterRow ecmID, alertID, customerName, counterparties, _
eventType, outputFile, toolVersion, notes, detail

Exit Sub
Fail:
' Never let audit logging itself break the calling macro's actual
' work - by the time this runs, the case file it's logging has
' already saved successfully. Silently swallow rather than risk a
' message-box storm from a shared low-level logger.
End Sub

'------------------------------------------------------------------
' Copies specific sheets from sourceWb (e.g. Module9's newWb, which
' already contains "Raw Transactions"/"Pivot"/"CP Selection"/
' "DeDupe") into this case's audit workbook, each renamed with a
' tag suffix - e.g. "Raw Transactions (Alerted)". Re-running the
' SAME tag replaces just those 4 tabs with the fresh version;
' running a DIFFERENT tag (e.g. Non-Alerted after Alerted) adds its
' own 4 tabs alongside rather than touching the first set.
'
'   sourceWb   - the workbook to copy FROM. Must already be open in
'                the current Excel session (this always opens the
'                case audit workbook in the SAME Application
'                instance, not a background ghost one, specifically
'                so this cross-workbook sheet copy is the ordinary,
'                well-trodden same-process kind).
'   sheetNames - array of sheet names to look for in sourceWb.
'                Missing ones (e.g. "Pivot" when there were no
'                counterparty rows to pivot) are silently skipped.
'   tagSuffix  - e.g. "Alerted" / "NonAlerted" - becomes part of
'                each destination tab's name.
'------------------------------------------------------------------
Public Sub ArchiveOutputSheets(ByVal ecmID As String, _
ByVal sourceWb As Workbook, _
ByVal sheetNames As Variant, _
ByVal tagSuffix As String)
On Error GoTo Fail
If Trim$(ecmID) = "" Then Exit Sub

Dim ghostApp As Object, wasAlreadyOpen As Boolean, isNewFile As Boolean
Dim wb As Workbook
Set wb = OpenCaseAuditWorkbook(ecmID, ghostApp, wasAlreadyOpen, isNewFile, useGhost:=False)
If wb Is Nothing Then Exit Sub

EnsureRegisterSheet wb   ' first-touch safety net if this case's very first event is a transaction consolidation

Dim nm As Variant, srcSheet As Worksheet, destName As String, oldWs As Worksheet
For Each nm In sheetNames
Set srcSheet = Nothing
On Error Resume Next
Set srcSheet = sourceWb.Sheets(CStr(nm))
On Error GoTo Fail

If Not srcSheet Is Nothing Then
destName = Left$(CStr(nm) & " (" & tagSuffix & ")", 31)

Set oldWs = Nothing
On Error Resume Next
Set oldWs = wb.Sheets(destName)
On Error GoTo Fail
If Not oldWs Is Nothing Then
wb.Application.DisplayAlerts = False
oldWs.Delete
wb.Application.DisplayAlerts = True
End If

srcSheet.Copy After:=wb.Sheets(wb.Sheets.Count)
wb.Sheets(wb.Sheets.Count).Name = destName
End If
Next nm

CloseCaseAuditWorkbook wb, ghostApp, wasAlreadyOpen, ecmID, isNewFile
Exit Sub

Fail:
On Error Resume Next
If Not ghostApp Is Nothing Then ghostApp.Quit
On Error GoTo 0
End Sub

'------------------------------------------------------------------
' Writes the full generated narrative/RFI text into a single
' "Narrative" tab in this case's audit workbook - always the latest
' version. If it's regenerated later, this tab is replaced outright;
' the Register tab's row history is what shows every time it was
' generated/regenerated and by whom, so nothing about *when* is
' lost even though only the latest text is kept here.
'
' Text is written one row per paragraph (not one giant cell) - far
' more readable in Excel, and avoids the 32,767-char cell cap for
' anything short of one freakishly long paragraph.
'------------------------------------------------------------------
Public Sub ArchiveNarrativeText(ByVal ecmID As String, ByVal fullText As String, ByVal docLabel As String)
On Error GoTo Fail
If Trim$(ecmID) = "" Then Exit Sub

Dim ghostApp As Object, wasAlreadyOpen As Boolean, isNewFile As Boolean
Dim wb As Workbook, ws As Worksheet

Set wb = OpenCaseAuditWorkbook(ecmID, ghostApp, wasAlreadyOpen, isNewFile)
If wb Is Nothing Then Exit Sub

EnsureRegisterSheet wb

On Error Resume Next
wb.Application.DisplayAlerts = False
wb.Sheets(NARRATIVE_SHEET_NAME).Delete
wb.Application.DisplayAlerts = True
On Error GoTo Fail

Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
ws.Name = NARRATIVE_SHEET_NAME
ws.Cells(1, 1).Value = "Generated"
ws.Cells(1, 2).Value = Now
ws.Cells(2, 1).Value = "Document"
ws.Cells(2, 2).Value = docLabel
ws.Cells(4, 1).Value = "Text"
ws.Cells(4, 1).Font.Bold = True

Dim lines() As String, i As Long
lines = Split(Replace(Replace(fullText, vbCrLf, vbCr), vbLf, vbCr), vbCr)
For i = LBound(lines) To UBound(lines)
ws.Cells(5 + i, 1).Value = lines(i)
Next i
ws.Columns("A").ColumnWidth = 120

CloseCaseAuditWorkbook wb, ghostApp, wasAlreadyOpen, ecmID, isNewFile
Exit Sub

Fail:
On Error Resume Next
If Not ghostApp Is Nothing Then ghostApp.Quit
On Error GoTo 0
End Sub

'==================================================================
' FIRM-WIDE (THIS WORKBOOK'S OWN) Audit_Ledger SHEET
'==================================================================

'------------------------------------------------------------------
' Finds the "Audit_Ledger" sheet, creating it (with headers) if
' missing. This name is deliberately distinct from any pre-existing
' "Audit_Log" tab a workbook might already have - this function
' never looks at, renames, or otherwise touches that sheet.
'------------------------------------------------------------------
Public Function GetOrCreateAuditSheet(ByVal wb As Workbook) As Worksheet
Dim ws As Worksheet
On Error Resume Next
Set ws = wb.Sheets(AUDIT_SHEET_NAME)
On Error GoTo 0

If ws Is Nothing Then
Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
ws.Name = AUDIT_SHEET_NAME
WriteAuditHeaders ws
End If

Set GetOrCreateAuditSheet = ws
End Function

Private Sub WriteAuditHeaders(ByVal ws As Worksheet)
ws.Cells(1, COL_TIMESTAMP).Value = "Date & Time"
ws.Cells(1, COL_ANALYST).Value = "Analyst"
ws.Cells(1, COL_ECMCASE).Value = "ECM Case ID"
ws.Cells(1, COL_ALERTID).Value = "Alert ID"
ws.Cells(1, COL_CUSTOMER).Value = "Customer Name"
ws.Cells(1, COL_COUNTERPARTIES).Value = "Counterparties"
ws.Cells(1, COL_EVENTTYPE).Value = "Event Type"
ws.Cells(1, COL_OUTPUTFILE).Value = "Output File"
ws.Cells(1, COL_TOOLVERSION).Value = "Tool Version"
ws.Cells(1, COL_NOTES).Value = "Notes"
ws.Cells(1, COL_DETAIL).Value = "Detail"
With ws.Range(ws.Cells(1, 1), ws.Cells(1, AUDIT_HEADER_COLS))
.Font.Bold = True
.Interior.Color = RGB(0, 70, 127)
.Font.Color = RGB(255, 255, 255)
End With
ws.Columns("A:K").AutoFit
End Sub

Private Sub WriteAuditRow(ByVal ws As Worksheet, _
ByVal ecmID As String, ByVal alertID As String, _
ByVal customerName As String, ByVal counterparties As String, _
ByVal eventType As String, ByVal outputFile As String, _
ByVal toolVersion As String, ByVal notes As String, ByVal detail As String)
Dim nextRow As Long
nextRow = ws.Cells(ws.Rows.Count, COL_TIMESTAMP).End(xlUp).Row + 1
ws.Cells(nextRow, COL_TIMESTAMP).Value = Now
ws.Cells(nextRow, COL_ANALYST).Value = Environ("USERNAME")
ws.Cells(nextRow, COL_ECMCASE).Value = ecmID
ws.Cells(nextRow, COL_ALERTID).Value = alertID
ws.Cells(nextRow, COL_CUSTOMER).Value = customerName
ws.Cells(nextRow, COL_COUNTERPARTIES).Value = counterparties
ws.Cells(nextRow, COL_EVENTTYPE).Value = eventType
ws.Cells(nextRow, COL_OUTPUTFILE).Value = outputFile
ws.Cells(nextRow, COL_TOOLVERSION).Value = toolVersion
ws.Cells(nextRow, COL_NOTES).Value = notes
ws.Cells(nextRow, COL_DETAIL).Value = detail
End Sub

'==================================================================
' PER-CASE AUDIT WORKBOOK (Desktop\{ECMID}\{ECMID}_Audit_Log.xlsx)
'==================================================================

Private Sub WriteCaseRegisterRow(ByVal ecmID As String, ByVal alertID As String, _
ByVal customerName As String, ByVal counterparties As String, _
ByVal eventType As String, ByVal outputFile As String, _
ByVal toolVersion As String, ByVal notes As String, ByVal detail As String)
On Error GoTo Fail
If Trim$(ecmID) = "" Then Exit Sub

Dim ghostApp As Object, wasAlreadyOpen As Boolean, isNewFile As Boolean
Dim wb As Workbook, ws As Worksheet

Set wb = OpenCaseAuditWorkbook(ecmID, ghostApp, wasAlreadyOpen, isNewFile)
If wb Is Nothing Then Exit Sub

Set ws = EnsureRegisterSheet(wb)
WriteAuditRow ws, ecmID, alertID, customerName, counterparties, _
eventType, outputFile, toolVersion, notes, detail

CloseCaseAuditWorkbook wb, ghostApp, wasAlreadyOpen, ecmID, isNewFile
Exit Sub

Fail:
On Error Resume Next
If Not ghostApp Is Nothing Then ghostApp.Quit
On Error GoTo 0
End Sub

' Gets or creates the "Register" tab. On a brand new case audit
' workbook, collapses Excel's default blank sheet(s) down to one and
' renames it, rather than leaving those sitting alongside it. Safe
' to call defensively from ArchiveOutputSheets/ArchiveNarrativeText
' too, in case a case's very first-ever event isn't an OSDD/folder-
' creation run that would normally create this tab first.
Private Function EnsureRegisterSheet(ByVal wb As Workbook) As Worksheet
Dim ws As Worksheet
On Error Resume Next
Set ws = wb.Sheets(REGISTER_SHEET_NAME)
On Error GoTo 0

If Not ws Is Nothing Then
Set EnsureRegisterSheet = ws
Exit Function
End If

Do While wb.Sheets.Count > 1
wb.Application.DisplayAlerts = False
wb.Sheets(wb.Sheets.Count).Delete
wb.Application.DisplayAlerts = True
Loop
Set ws = wb.Sheets(1)
ws.Name = REGISTER_SHEET_NAME
WriteAuditHeaders ws
Set EnsureRegisterSheet = ws
End Function

' Opens (or creates) this case's Desktop\{ECMID}\{ECMID}_Audit_Log.xlsx.
'
'   useGhost = True (default) - opens via a separate, invisible
'       Excel.Application, for callers that only write cell values/
'       text (WriteCaseRegisterRow, ArchiveNarrativeText). Silent,
'       no window flicker.
'   useGhost = False - opens via the CURRENT Application instead, so
'       a caller that needs to Worksheet.Copy a sheet IN from
'       another workbook (ArchiveOutputSheets) does an ordinary
'       same-process cross-workbook copy - the well-trodden,
'       low-risk path - rather than a copy across two separate
'       Excel processes.
'
' wasAlreadyOpen/isNewFile are ByRef so CloseCaseAuditWorkbook knows
' how to save/close it correctly afterward.
Private Function OpenCaseAuditWorkbook(ByVal ecmID As String, _
ByRef ghostApp As Object, _
ByRef wasAlreadyOpen As Boolean, _
ByRef isNewFile As Boolean, _
Optional ByVal useGhost As Boolean = True) As Workbook
On Error GoTo Fail

Dim xlsxPath As String, xlsxName As String
xlsxPath = GetCaseAuditXlsxPath(ecmID)
xlsxName = Mid$(xlsxPath, InStrRev(xlsxPath, "\") + 1)

wasAlreadyOpen = False
isNewFile = False
Set ghostApp = Nothing

Dim pushWb As Workbook
For Each pushWb In Application.Workbooks
If pushWb.Name = xlsxName Then
wasAlreadyOpen = True
Set OpenCaseAuditWorkbook = pushWb
UnprotectCaseWorkbook OpenCaseAuditWorkbook
Exit Function
End If
Next pushWb

If useGhost Then
Set ghostApp = CreateObject("Excel.Application")
ghostApp.Visible = False
ghostApp.DisplayAlerts = False
ghostApp.EnableEvents = False

If Dir(xlsxPath) <> "" Then
Set OpenCaseAuditWorkbook = ghostApp.Workbooks.Open(fileName:=xlsxPath, UpdateLinks:=False)
Else
isNewFile = True
Set OpenCaseAuditWorkbook = ghostApp.Workbooks.Add
End If
Else
If Dir(xlsxPath) <> "" Then
Set OpenCaseAuditWorkbook = Workbooks.Open(fileName:=xlsxPath, UpdateLinks:=False)
Else
isNewFile = True
Set OpenCaseAuditWorkbook = Workbooks.Add
End If
End If
UnprotectCaseWorkbook OpenCaseAuditWorkbook
Exit Function

Fail:
On Error Resume Next
If Not ghostApp Is Nothing Then ghostApp.Quit
Set ghostApp = Nothing
On Error GoTo 0
Set OpenCaseAuditWorkbook = Nothing
End Function

Private Sub CloseCaseAuditWorkbook(ByVal wb As Workbook, ByVal ghostApp As Object, _
ByVal wasAlreadyOpen As Boolean, ByVal ecmID As String, ByVal isNewFile As Boolean)
On Error Resume Next

Dim origAlerts As Boolean
origAlerts = wb.Application.DisplayAlerts
wb.Application.DisplayAlerts = False

' Lock everything back down right before it lands on disk - the
' file is meant to be read-only to anyone opening it directly;
' only this module's own code (which knows the password) can
' unprotect and update it on the next event.
ProtectCaseWorkbook wb

If isNewFile Then
wb.SaveAs fileName:=GetCaseAuditXlsxPath(ecmID), FileFormat:=51
Else
wb.Save
End If

wb.Application.DisplayAlerts = origAlerts

If Not wasAlreadyOpen Then
wb.Close SaveChanges:=False   ' already saved above
If Not ghostApp Is Nothing Then
ghostApp.Quit
Set ghostApp = Nothing
End If
End If

On Error GoTo 0
End Sub

' Removes sheet + workbook-structure protection so this run's code
' can freely add/rename/delete/edit sheets. Harmless no-op on a
' brand-new, never-yet-protected workbook. Always called right
' after opening; always re-locked by ProtectCaseWorkbook before the
' next save.
Private Sub UnprotectCaseWorkbook(ByVal wb As Workbook)
On Error Resume Next
wb.Unprotect Password:=AUDIT_PROTECT_PASSWORD
Dim ws As Worksheet
For Each ws In wb.Sheets
ws.Unprotect Password:=AUDIT_PROTECT_PASSWORD
Next ws
On Error GoTo 0
End Sub

' Re-locks every sheet (cells can't be edited without the password)
' and the workbook structure (sheets can't be added/renamed/deleted/
' unhidden without the password). Called right before every save,
' from CloseCaseAuditWorkbook, so the file always lands on disk
' read-only to anyone opening it without AUDIT_PROTECT_PASSWORD.
Private Sub ProtectCaseWorkbook(ByVal wb As Workbook)
On Error Resume Next
Dim ws As Worksheet
For Each ws In wb.Sheets
ws.Protect Password:=AUDIT_PROTECT_PASSWORD
Next ws
wb.Protect Password:=AUDIT_PROTECT_PASSWORD, Structure:=True, Windows:=False
On Error GoTo 0
End Sub

Private Function GetCaseDesktopFolder(ByVal ecmID As String) As String
Dim FSO As Object, desktopPath As String, userProfile As String
Set FSO = CreateObject("Scripting.FileSystemObject")
userProfile = Environ("USERPROFILE")
If FSO.FolderExists(userProfile & "\OneDrive - Community Federal Savings Bank\Desktop") Then
desktopPath = userProfile & "\OneDrive - Community Federal Savings Bank\Desktop"
Else
desktopPath = userProfile & "\Desktop"
End If

Dim folderPath As String
folderPath = desktopPath & "\" & ecmID
If Not FSO.FolderExists(folderPath) Then FSO.CreateFolder folderPath
GetCaseDesktopFolder = folderPath
End Function

Private Function GetCaseAuditXlsxPath(ByVal ecmID As String) As String
GetCaseAuditXlsxPath = GetCaseDesktopFolder(ecmID) & "\" & ecmID & AUDIT_XLSX_SUFFIX
End Function

'------------------------------------------------------------------
' Shared helper: builds the semicolon-joined counterparty name list
' from Sheet1!J19:J24, the same range every macro in this workbook
' already reads counterparties from. Centralized here so every
' LogAuditEvent call gets the same list without each caller
' re-implementing the loop.
'------------------------------------------------------------------
Public Function GetCounterpartyList(ByVal ws As Worksheet) As String
Dim i As Long, cpName As String, list As String
For i = 19 To 24
cpName = Trim$(CStr(ws.Range("J" & i).Value))
If cpName <> "" Then
If list <> "" Then list = list & "; "
list = list & cpName
End If
Next i
GetCounterpartyList = list
End Function
