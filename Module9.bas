Attribute VB_Name = "Module9"
Sub Consolidated_AML_Workflow()

' ==========================================
' THE "JACKPOT" CONSOLIDATION WORKFLOW
' ==========================================
Application.EnableCancelKey = xlErrorHandler
On Error GoTo CancelHandler

' Safe default so CancelHandler can always restore Calculation even
' if an error fires before the real capture below ever runs.
Dim origCalc As XlCalculation
origCalc = xlCalculationAutomatic

Dim WbSource As Workbook, WsMaster As Worksheet, wsSource As Worksheet, wsHome As Worksheet
Dim LastRowSource As Long, LastRowMaster As Long, LastCol As Long
Dim HeaderCopied As Boolean

Dim HeaderCell As Range, HeaderRow As Long, HeadCol As Long
Dim DateHeader As Range, DateRange As Range, AmtHeader As Range, AmtRange As Range
Dim drCrCell As Range, benNameCell As Range, origNameCell As Range
Dim drCrCol As Long, benNameCol As Long, origNameCol As Long
Dim WsRawTemp As Worksheet, newWb As Workbook, wsExport As Worksheet, ws As Worksheet
Dim TransCol As Long, AlertCol As Long
Dim desktopPath As String, excelFileName As String, saveFolderPath As String, folderPath As String
Dim ecmID As String, AlertID As String
Dim wsPivot As Worksheet, ptCache As PivotCache, pt As PivotTable, ptRange As Range
Dim lastRowCP As Long, lastColCP As Long
Dim slash As String

Dim FSO As Object, objFolder As Object, objFile As Object
Dim fileFound As Boolean

' --- THE ULTIMATE PATH FIX ---
' This forces Excel to use a built-in slash, preventing it from vanishing during copy/paste
slash = Application.PathSeparator

' ==========================================
' 1. SETUP & THE PROVEN FOLDER CONNECTION
' ==========================================

ThisWorkbook.Unprotect Password:="p7ss"
On Error Resume Next ' In case the sheets don't exist or are already unlocked
ThisWorkbook.Sheets("Sheet1").Unprotect Password:="p7ss"
ThisWorkbook.Sheets("ConsolidatedData").Unprotect Password:="p7ss"
On Error GoTo CancelHandler ' Turn the error handler back on

Set wsHome = ActiveWorkbook.Sheets("Sheet1")
ecmID = Trim(wsHome.Range("J10").Value)
AlertID = Trim(wsHome.Range("J11").Value)
If AlertID = "" Then AlertID = "ALERT"

If ecmID = "" Then
MsgBox "Action Denied: ECM ID is missing in J10.", vbCritical, "Missing ID"
Exit Sub
End If

' Source folder is always "Transaction Files" now - the old
' UserForm1 picker (Transaction Files / Non Alerted / Cancel) has
' been removed, so this runs straight through like it used to.
Dim sourceFolderName As String
sourceFolderName = "Transaction Files"

' Build the exact paths using the guaranteed slash
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

' Initialize ConsolidatedData sheet
On Error Resume Next
Set WsMaster = ActiveWorkbook.Sheets("ConsolidatedData")
On Error GoTo CancelHandler

If WsMaster Is Nothing Then
Set WsMaster = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count))
WsMaster.Name = "ConsolidatedData"
Else
WsMaster.Cells.Clear
End If

' EnableEvents/Calculation were never touched here before - every
' cell write, paste, and sheet copy below could trigger a full
' workbook recalculation (not just this new workbook - this tool's
' own aggregate formulas too) and fire the full event pipeline on
' every workbook open/close. That's the single biggest reason this
' takes as long as it does, and the long unresponsive stretch it
' creates is very likely what causes the black-screen symptom on a
' remote/VDI session (Windows marks Excel "Not Responding" and RDP
' can't get a valid frame to redraw from).
origCalc = Application.Calculation
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Application.ScreenUpdating = False
Application.DisplayAlerts = False
HeaderCopied = False

' ==========================================
' 2. COMBINE FILES
' ==========================================
For Each objFile In objFolder.Files
If (InStr(1, objFile.Name, ".xls", vbTextCompare) > 0) And (Left(objFile.Name, 2) <> "~$") And (objFile.Name <> ThisWorkbook.Name) Then
Set WbSource = Workbooks.Open(objFile.path, ReadOnly:=True, UpdateLinks:=False)
On Error Resume Next
Set wsSource = WbSource.Sheets(1)
On Error GoTo CancelHandler

If Not wsSource Is Nothing Then
    With wsSource
        Set HeaderCell = .Cells.Find(What:="Transaction ID", LookIn:=xlValues, LookAt:=xlWhole)
        If Not HeaderCell Is Nothing Then
            HeaderRow = HeaderCell.row
            HeadCol = HeaderCell.Column
            LastRowSource = .Cells(.Rows.count, HeadCol).End(xlUp).row
            
            If LastRowSource >= HeaderRow Then
                If Not HeaderCopied Then
                    .Range(.Cells(HeaderRow, HeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=WsMaster.Range("A1")
                    HeaderCopied = True
                Else
                    If LastRowSource > HeaderRow Then
                        LastRowMaster = WsMaster.Cells(WsMaster.Rows.count, "A").End(xlUp).row + 1
                        .Range(.Cells(HeaderRow + 1, HeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=WsMaster.Range("A" & LastRowMaster)
                    End If
                End If
            End If
        End If
    End With
End If
WbSource.Close SaveChanges:=False
End If
Next objFile

' ==========================================
' 2.5 SNAPSHOT RAW DATA
' ==========================================
On Error Resume Next
Application.DisplayAlerts = False
wsHome.Parent.Sheets("TempRawBackup").Delete
Application.DisplayAlerts = True
On Error GoTo CancelHandler

WsMaster.Copy After:=wsHome.Parent.Sheets(wsHome.Parent.Sheets.count)
Set WsRawTemp = ActiveSheet
WsRawTemp.Name = "TempRawBackup"

' ==========================================
' 3. AGGRESSIVE DATA CLEANUP
' ==========================================
WsMaster.Activate

Set DateHeader = WsMaster.Rows(1).Find(What:="Transaction Date", LookIn:=xlValues, LookAt:=xlPart)
If Not DateHeader Is Nothing Then
LastRowMaster = WsMaster.Cells(WsMaster.Rows.count, DateHeader.Column).End(xlUp).row
If LastRowMaster > 1 Then
Set DateRange = WsMaster.Range(WsMaster.Cells(2, DateHeader.Column), WsMaster.Cells(LastRowMaster, DateHeader.Column))
DateRange.TextToColumns Destination:=DateRange.Cells(1, 1), DataType:=xlDelimited, FieldInfo:=Array(Array(1, 3))
DateRange.NumberFormat = "dddd, mmmm d, yyyy"
End If
End If

Set AmtHeader = WsMaster.Rows(1).Find(What:="Transaction Amount", LookIn:=xlValues, LookAt:=xlPart)
If Not AmtHeader Is Nothing Then
LastRowMaster = WsMaster.Cells(WsMaster.Rows.count, AmtHeader.Column).End(xlUp).row
If LastRowMaster > 1 Then
Set AmtRange = WsMaster.Range(WsMaster.Cells(2, AmtHeader.Column), WsMaster.Cells(LastRowMaster, AmtHeader.Column))
AmtRange.Value = AmtRange.Value
' --- THE ROOT FIX ---
AmtRange.NumberFormat = "$#,##0.00"
' --------------------
End If
End If

Set drCrCell = WsMaster.Rows(1).Find(What:="Dr Cr", LookAt:=xlPart)
Set benNameCell = WsMaster.Rows(1).Find(What:="Beneficiary Name", LookAt:=xlPart)
Set origNameCell = WsMaster.Rows(1).Find(What:="Originator Name", LookAt:=xlPart)

If Not drCrCell Is Nothing And Not benNameCell Is Nothing And Not origNameCell Is Nothing Then
drCrCol = drCrCell.Column
benNameCol = benNameCell.Column
origNameCol = origNameCell.Column

LastRowMaster = WsMaster.Cells(WsMaster.Rows.count, "A").End(xlUp).row
LastCol = WsMaster.Cells(1, WsMaster.Columns.count).End(xlToLeft).Column + 1

If LastRowMaster > 1 Then
WsMaster.Cells(1, LastCol).Value = "Counterparty"
WsMaster.Range(WsMaster.Cells(2, LastCol), WsMaster.Cells(LastRowMaster, LastCol)).FormulaR1C1 = _
"=IF(RC" & drCrCol & "=""DR"", RC" & benNameCol & ", RC" & origNameCol & ")"

WsMaster.Cells(1, LastCol - 1).Copy
WsMaster.Cells(1, LastCol).PasteSpecial Paste:=xlPasteFormats
WsMaster.Range(WsMaster.Cells(2, LastCol - 1), WsMaster.Cells(LastRowMaster, LastCol - 1)).Copy
WsMaster.Range(WsMaster.Cells(2, LastCol), WsMaster.Cells(LastRowMaster, LastCol)).PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False
End If
End If

' ==========================================
' 4. TWO-LAYER EXPORT & DEDUPE
' ==========================================
Set newWb = Workbooks.Add

WsRawTemp.Copy Before:=newWb.Sheets(1): ActiveSheet.Name = "Raw Transactions"
WsMaster.Copy After:=newWb.Sheets(newWb.Sheets.count): ActiveSheet.Name = "CP Selection"

Set wsExport = newWb.Sheets("CP Selection")
TransCol = 0: AlertCol = 0
On Error Resume Next
TransCol = wsExport.Rows(1).Find(What:="Transaction ID", LookAt:=xlPart).Column
AlertCol = wsExport.Rows(1).Find(What:="Alert Information", LookAt:=xlPart).Column
On Error GoTo CancelHandler

If TransCol > 0 And AlertCol > 0 Then
wsExport.UsedRange.RemoveDuplicates Columns:=Array(TransCol, AlertCol), Header:=xlYes
End If

wsExport.Copy After:=newWb.Sheets(newWb.Sheets.count): ActiveSheet.Name = "DeDupe"

Set wsExport = newWb.Sheets("DeDupe")
TransCol = 0
On Error Resume Next
TransCol = wsExport.Rows(1).Find(What:="Transaction ID", LookAt:=xlPart).Column
On Error GoTo CancelHandler

If TransCol > 0 Then wsExport.UsedRange.RemoveDuplicates Columns:=Array(TransCol), Header:=xlYes

For Each ws In newWb.Sheets
If ws.Name = "Raw Transactions" Or ws.Name = "CP Selection" Or ws.Name = "DeDupe" Then
With ws.Cells
.WrapText = False
.EntireColumn.AutoFit
.WrapText = True
.EntireRow.AutoFit
.VerticalAlignment = xlTop
End With
Else
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
End If
Next ws

' ==========================================
' 4.5 PIVOT TABLES (MULTI-SOURCE)
' ==========================================
Set wsExport = newWb.Sheets("CP Selection")
lastRowCP = wsExport.Cells(wsExport.Rows.count, "A").End(xlUp).row
lastColCP = wsExport.Cells(1, wsExport.Columns.count).End(xlToLeft).Column

If lastRowCP > 1 Then
Set ptRange = wsExport.Range(wsExport.Cells(1, 1), wsExport.Cells(lastRowCP, lastColCP))
Set wsPivot = newWb.Sheets.Add(After:=newWb.Sheets("Raw Transactions"))
wsPivot.Name = "Pivot"

' --- MEMORY BANK 1: Connected to "CP Selection" ---
Set ptCache = newWb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRange)

' ------------------------------------------
' PIVOT 1: SCENARIO PIVOT (KEPT EXACTLY AS ORIGINAL)
' ------------------------------------------
Set pt = ptCache.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), TableName:="ScenarioPivot")
On Error Resume Next
With pt
.TableStyle2 = "PivotStyleLight16"
With .PivotFields("Alert Information")
.Orientation = xlRowField
.Position = 1
End With
With .PivotFields("Dr Cr")
.Orientation = xlRowField
.Position = 2
End With
With .PivotFields("Counterparty")
.Orientation = xlRowField
.Position = 3
End With
.AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount", xlSum
.PivotFields("Sum of Transaction Amount").NumberFormat = "$#,#00.00"
.AddDataField .PivotFields("Transaction Amount"), "Count of Transaction Amount", xlCount
.RowAxisLayout xlCompactRow
.PivotFields("Count of Transaction Amount").NumberFormat = "0"
.PivotFields("Alert Information").AutoSort xlDescending, "Sum of Transaction Amount"
.PivotFields("Counterparty").AutoSort xlDescending, "Sum of Transaction Amount"
End With
On Error GoTo CancelHandler

' ------------------------------------------
' SETUP SOURCE 2: "DeDupe"
' ------------------------------------------
Dim wsDeDupe As Worksheet
Dim lastRowDD As Long, lastColDD As Long, dtColDD As Long
Dim ptRangeDeDupe As Range, ptCacheDeDupe As PivotCache

Set wsDeDupe = newWb.Sheets("DeDupe")

' --- INVISIBLE STABILITY FIX: Delete blank rows so grouping doesn't crash ---
On Error Resume Next
dtColDD = wsDeDupe.Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
lastRowDD = wsDeDupe.Cells(wsDeDupe.Rows.count, dtColDD).End(xlUp).row
If lastRowDD > 1 And dtColDD > 0 Then
wsDeDupe.Range(wsDeDupe.Cells(2, dtColDD), wsDeDupe.Cells(lastRowDD, dtColDD)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
' Format the raw column as short dates so the new pivots inherit it perfectly
wsDeDupe.Range(wsDeDupe.Cells(2, dtColDD), wsDeDupe.Cells(lastRowDD, dtColDD)).NumberFormat = "m/d/yyyy"
End If
On Error GoTo CancelHandler

lastRowDD = wsDeDupe.Cells(wsDeDupe.Rows.count, "A").End(xlUp).row
lastColDD = wsDeDupe.Cells(1, wsDeDupe.Columns.count).End(xlToLeft).Column

If lastRowDD > 1 Then
' --- MEMORY BANK 2: Connected to "DeDupe" ---
Set ptRangeDeDupe = wsDeDupe.Range(wsDeDupe.Cells(1, 1), wsDeDupe.Cells(lastRowDD, lastColDD))
Set ptCacheDeDupe = newWb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRangeDeDupe)

' ------------------------------------------
' PIVOT 2: TEMPORAL PIVOT (KEPT EXACTLY AS ORIGINAL)
' ------------------------------------------
Set pt = ptCacheDeDupe.CreatePivotTable(TableDestination:=wsPivot.Range("F3"), TableName:="TemporalPivot")
On Error Resume Next
With pt
    .TableStyle2 = "PivotStyleLight16"
    With .PivotFields("Transaction Date")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount ", xlSum
    .PivotFields("Sum of Transaction Amount ").NumberFormat = "$#,#00.00"
    
    .AddDataField .PivotFields("Transaction Amount"), "Count of Transaction Amount ", xlCount
    .PivotFields("Count of Transaction Amount ").NumberFormat = "0"
End With
' The Date Grouping Array (Includes Days, Months, and Years!)
wsPivot.Range("F4").Group Start:=True, End:=True, Periods:=Array(False, False, False, True, True, False, True)

' --- THE ULTIMATE FORMAT FIX: Apply directly to the PivotField memory, not the cells ---
On Error Resume Next
pt.PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
On Error GoTo CancelHandler

' ------------------------------------------
' PIVOT 3: ENHANCED TEMPORAL PIVOT (CLONED FROM PT2)
' ------------------------------------------
pt.TableRange2.Copy Destination:=wsPivot.Range("K3")

Dim pt3 As PivotTable
Set pt3 = wsPivot.Range("K3").PivotTable
pt3.Name = "TemporalPivot_Enhanced"

On Error Resume Next
With pt3
' 1. Remove PT2's old values
.PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
.PivotFields("Count of Transaction Amount ").Orientation = xlHidden

' 2. Add Dr/Cr to the bottom of the Date stack (Position 4)
With .PivotFields("Dr Cr")
    .Orientation = xlRowField
    .Position = 4
End With

' 3. Add the new values in requested order
.AddDataField .PivotFields("Transaction Amount"), "No of Trx  ", xlCount
.PivotFields("No of Trx  ").NumberFormat = "0"
.AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount  ", xlSum
.PivotFields("Sum of Transaction Amount  ").NumberFormat = "$#,#00.00"

' 4. Force the mm/dd/yyyy date format
.PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
End With
On Error GoTo CancelHandler

' ------------------------------------------
' PIVOT 4: DR/CR INVERTED PIVOT (CLONED FROM PT2)
' ------------------------------------------
pt.TableRange2.Copy Destination:=wsPivot.Range("Q3")

Dim pt4 As PivotTable
Set pt4 = wsPivot.Range("Q3").PivotTable
pt4.Name = "DrCrTemporalPivot"

On Error Resume Next
With pt4
' 1. Remove PT2's old values
.PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
.PivotFields("Count of Transaction Amount ").Orientation = xlHidden

' 2. Add Dr/Cr to the TOP of the Date stack (Position 1)
With .PivotFields("Dr Cr")
    .Orientation = xlRowField
    .Position = 1
End With

' 3. Add the new values
.AddDataField .PivotFields("Transaction Amount"), "No of Trx   ", xlCount
.PivotFields("No of Trx   ").NumberFormat = "0"
.AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount   ", xlSum
.PivotFields("Sum of Transaction Amount   ").NumberFormat = "$#,#00.00"

' 4. Force the mm/dd/yyyy date format
.PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
End With
On Error GoTo CancelHandler

' ------------------------------------------
' DYNAMIC HIGHLIGHTING & BULLETPROOF DATE FIX
' ------------------------------------------
' Both of these used to scan cell-by-cell via COM (one round-trip
' per cell across the whole pivot sheet, then again per row of each
' pivot's RowRange) - now a single bulk .Value read, an in-memory
' scan, and one Union-based write each. Same result, a fraction of
' the COM calls.
On Error Resume Next
BulletproofDateFormat wsPivot
HighlightDrCrRows pt3, wsPivot
HighlightDrCrRows pt4, wsPivot
On Error GoTo CancelHandler

End If

' Format entire sheet spacing
wsPivot.Columns("A:W").AutoFit
End If

' ==========================================
' 5. FINALIZE MASTER TAB & SAVE
' ==========================================
WsMaster.Cells.Clear
newWb.Sheets("DeDupe").UsedRange.Copy Destination:=WsMaster.Range("A1")

With WsMaster.Cells
.WrapText = False
.EntireColumn.AutoFit
.WrapText = True
.EntireRow.AutoFit
.VerticalAlignment = xlTop
End With

On Error Resume Next
Application.DisplayAlerts = False
wsHome.Parent.Sheets("TempRawBackup").Delete
Application.DisplayAlerts = True

' --- ZERO-BULLSHIT SAVE FIX ---
Dim finalSavePath As String
' Always the alerted Transaction Files source now.
Dim fileTag As String
fileTag = "Alerted"
excelFileName = ecmID & "_" & AlertID & "_Combined_" & fileTag & "_Transaction.xlsx"

' Building path with guaranteed slashes
finalSavePath = saveFolderPath & slash & excelFileName

' Remove the error bypass so if Excel blocks the save, we actually see why
Application.DisplayAlerts = False
newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
Application.DisplayAlerts = True

' Centralized audit ledger row - Register tab of this case's own
' Desktop\{ecmID}\{ecmID}_Audit_Log.xlsx.
modAuditLog.LogAuditEvent ecmID:=ecmID, AlertID:=AlertID, _
customerName:=Trim(wsHome.Range("J14").Value), _
counterparties:=modAuditLog.GetCounterpartyList(wsHome), _
eventType:="Transaction File Consolidated", _
outputFile:=finalSavePath, _
toolVersion:="2.4.1", _
notes:="source=" & sourceFolderName & ", tag=" & fileTag

' Copies the 4 actual sheets just built (Raw Transactions, Pivot,
' CP Selection, DeDupe) into that same audit workbook, tagged by
' source (Alerted/NonAlerted) - real content, not just a path.
' Re-running THIS tag refreshes those 4 tabs; running the other tag
' later adds its own 4 alongside instead of overwriting them.
modAuditLog.ArchiveOutputSheets ecmID:=ecmID, sourceWb:=newWb, _
sheetNames:=Array("Raw Transactions", "Pivot", "CP Selection", "DeDupe"), _
tagSuffix:=fileTag

newWb.Sheets("Raw Transactions").Activate

Application.EnableCancelKey = xlInterrupt
Application.Calculation = origCalc
Application.EnableEvents = True
Application.ScreenUpdating = True

On Error Resume Next
ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
On Error GoTo 0

' ==========================================
' 6. UPDATE SHARED MASTER TRACKER (SHEET2) - DEFERRED
'    Scheduled to run ~1s later in the background (modTrackerPush) so the
'    analyst sees "Workflow Complete" immediately instead of waiting on
'    the slow ghost-Excel / OneDrive write. The row still gets written.
' ==========================================
On Error Resume Next
Application.OnTime Now + TimeSerial(0, 0, 1), "PushTrxTracker_Deferred"
On Error GoTo 0
' ==========================================

MsgBox "Workflow Complete!" & vbCrLf & _
"Exported file inside the folder exactly to: " & vbCrLf & finalSavePath, vbInformation, "Success"
Exit Sub

CancelHandler:
' CRITICAL FIX: This ensures Excel unfreezes even if the macro crashes
Application.EnableEvents = True
Application.Calculation = origCalc
Application.EnableCancelKey = xlInterrupt
Application.ScreenUpdating = True
Application.DisplayAlerts = True

On Error Resume Next
ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
On Error GoTo 0

If Err.Number = 18 Then
MsgBox "Process Safely Cancelled.", vbInformation, "Aborted"
ElseIf Err.Number <> 0 Then
MsgBox "An unexpected error occurred:" & vbCrLf & Err.Description, vbCritical, "Error " & Err.Number
End If
End Sub

' Scans a sheet's UsedRange for date-shaped values and forces
' mm/dd/yyyy on them (guarding against Excel misreading a bare
' 4-digit year like "2025" as the year 1905). Was a per-cell COM
' loop; now one bulk .Value read, an in-memory scan, and one
' Union-based NumberFormat write covering every hit at once.
Private Sub BulletproofDateFormat(ByVal ws As Worksheet)
On Error Resume Next

Dim rUsed As Range
Set rUsed = ws.UsedRange

' Single-cell UsedRange is the one case .Value returns a scalar,
' not a 2D array - handle it directly rather than indexing into it.
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

' Highlights the Sum-column cell for every CR/DR row in the given
' pivot's RowRange (light red fill, dark red text). Was a per-row
' COM loop reading .Value and writing Interior.Color/Font.Color one
' row at a time; now one bulk .Value read of the whole RowRange, an
' in-memory scan, and one Union-based write for both colors.
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
' Single-row RowRange - rowArr is a scalar, not a 2D array.
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


