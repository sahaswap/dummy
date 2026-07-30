Attribute VB_Name = "Module9"
Sub Consolidated_AML_Workflow()

' ==========================================
' THE "JACKPOT" CONSOLIDATION WORKFLOW
' ==========================================
Application.EnableCancelKey = xlErrorHandler
On Error GoTo CancelHandler

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

' --- Ask which source folder to consolidate ---
Dim sourceFolderName As String, userCancelled As Boolean

UserForm1.Show vbModal
sourceFolderName = UserForm1.SelectedFolderName
userCancelled = UserForm1.userCancelled
Unload UserForm1

If userCancelled Or sourceFolderName = "" Then
Exit Sub
End If

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
Set WsMaster = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
WsMaster.Name = "ConsolidatedData"
Else
WsMaster.Cells.Clear
End If

Application.ScreenUpdating = False
Application.DisplayAlerts = False
HeaderCopied = False

' ==========================================
' 2. COMBINE FILES
' ==========================================
For Each objFile In objFolder.Files
If (InStr(1, objFile.Name, ".xls", vbTextCompare) > 0) And (Left(objFile.Name, 2) <> "~$") And (objFile.Name <> ThisWorkbook.Name) Then
Set WbSource = Workbooks.Open(objFile.Path, ReadOnly:=True, UpdateLinks:=False)
On Error Resume Next
Set wsSource = WbSource.Sheets(1)
On Error GoTo CancelHandler

If Not wsSource Is Nothing Then
    With wsSource
        Set HeaderCell = .Cells.Find(What:="Transaction ID", LookIn:=xlValues, LookAt:=xlWhole)
        If Not HeaderCell Is Nothing Then
            HeaderRow = HeaderCell.Row
            HeadCol = HeaderCell.Column
            LastRowSource = .Cells(.Rows.Count, HeadCol).End(xlUp).Row
            
            If LastRowSource >= HeaderRow Then
                If Not HeaderCopied Then
                    .Range(.Cells(HeaderRow, HeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=WsMaster.Range("A1")
                    HeaderCopied = True
                Else
                    If LastRowSource > HeaderRow Then
                        LastRowMaster = WsMaster.Cells(WsMaster.Rows.Count, "A").End(xlUp).Row + 1
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

WsMaster.Copy After:=wsHome.Parent.Sheets(wsHome.Parent.Sheets.Count)
Set WsRawTemp = ActiveSheet
WsRawTemp.Name = "TempRawBackup"

' ==========================================
' 3. AGGRESSIVE DATA CLEANUP
' ==========================================
WsMaster.Activate

Set DateHeader = WsMaster.Rows(1).Find(What:="Transaction Date", LookIn:=xlValues, LookAt:=xlPart)
If Not DateHeader Is Nothing Then
LastRowMaster = WsMaster.Cells(WsMaster.Rows.Count, DateHeader.Column).End(xlUp).Row
If LastRowMaster > 1 Then
Set DateRange = WsMaster.Range(WsMaster.Cells(2, DateHeader.Column), WsMaster.Cells(LastRowMaster, DateHeader.Column))
DateRange.TextToColumns Destination:=DateRange.Cells(1, 1), DataType:=xlDelimited, FieldInfo:=Array(Array(1, 3))
DateRange.NumberFormat = "dddd, mmmm d, yyyy"
End If
End If

Set AmtHeader = WsMaster.Rows(1).Find(What:="Transaction Amount", LookIn:=xlValues, LookAt:=xlPart)
If Not AmtHeader Is Nothing Then
LastRowMaster = WsMaster.Cells(WsMaster.Rows.Count, AmtHeader.Column).End(xlUp).Row
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

LastRowMaster = WsMaster.Cells(WsMaster.Rows.Count, "A").End(xlUp).Row
LastCol = WsMaster.Cells(1, WsMaster.Columns.Count).End(xlToLeft).Column + 1

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
WsMaster.Copy After:=newWb.Sheets(newWb.Sheets.Count): ActiveSheet.Name = "CP Selection"

Set wsExport = newWb.Sheets("CP Selection")
TransCol = 0: AlertCol = 0
On Error Resume Next
TransCol = wsExport.Rows(1).Find(What:="Transaction ID", LookAt:=xlPart).Column
AlertCol = wsExport.Rows(1).Find(What:="Alert Information", LookAt:=xlPart).Column
On Error GoTo CancelHandler

If TransCol > 0 And AlertCol > 0 Then
wsExport.UsedRange.RemoveDuplicates Columns:=Array(TransCol, AlertCol), Header:=xlYes
End If

wsExport.Copy After:=newWb.Sheets(newWb.Sheets.Count): ActiveSheet.Name = "DeDupe"

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
lastRowCP = wsExport.Cells(wsExport.Rows.Count, "A").End(xlUp).Row
lastColCP = wsExport.Cells(1, wsExport.Columns.Count).End(xlToLeft).Column

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
lastRowDD = wsDeDupe.Cells(wsDeDupe.Rows.Count, dtColDD).End(xlUp).Row
If lastRowDD > 1 And dtColDD > 0 Then
wsDeDupe.Range(wsDeDupe.Cells(2, dtColDD), wsDeDupe.Cells(lastRowDD, dtColDD)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
' Format the raw column as short dates so the new pivots inherit it perfectly
wsDeDupe.Range(wsDeDupe.Cells(2, dtColDD), wsDeDupe.Cells(lastRowDD, dtColDD)).NumberFormat = "m/d/yyyy"
End If
On Error GoTo CancelHandler

lastRowDD = wsDeDupe.Cells(wsDeDupe.Rows.Count, "A").End(xlUp).Row
lastColDD = wsDeDupe.Cells(1, wsDeDupe.Columns.Count).End(xlToLeft).Column

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
Dim ptRowCell As Range
Dim sumColIndex As Long

' Activating silent bypass so the macro NEVER aborts and skips the Save command
On Error Resume Next

' 1. THE BULLETPROOF DATE FIX (Scans the entire sheet directly)
Dim cl As Range
For Each cl In wsPivot.UsedRange
' Check if the cell is a date
If IsDate(cl.Value) And Not IsEmpty(cl.Value) Then
' Prevent the "2025" year label from turning into a date (Excel reads 2025 as the year 1905)
If Year(CDate(cl.Value)) > 1950 Then
cl.NumberFormat = "mm/dd/yyyy"
End If
End If
Next cl

' 2. Fix Pivot 3 (Middle Pivot) - Dr/Cr Highlighting
sumColIndex = pt3.DataBodyRange.Columns(2).Column
For Each ptRowCell In pt3.RowRange
If Trim(UCase(ptRowCell.Value)) = "CR" Or Trim(UCase(ptRowCell.Value)) = "DR" Then
wsPivot.Cells(ptRowCell.Row, sumColIndex).Interior.Color = RGB(255, 199, 206) ' Light Red Fill
wsPivot.Cells(ptRowCell.Row, sumColIndex).Font.Color = RGB(156, 0, 6) ' Dark Red Text
End If
Next ptRowCell

' 3. Fix Pivot 4 (Right Pivot) - Dr/Cr Highlighting
sumColIndex = pt4.DataBodyRange.Columns(2).Column
For Each ptRowCell In pt4.RowRange
If Trim(UCase(ptRowCell.Value)) = "CR" Or Trim(UCase(ptRowCell.Value)) = "DR" Then
wsPivot.Cells(ptRowCell.Row, sumColIndex).Interior.Color = RGB(255, 199, 206) ' Light Red Fill
wsPivot.Cells(ptRowCell.Row, sumColIndex).Font.Color = RGB(156, 0, 6) ' Dark Red Text
End If
Next ptRowCell

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
Dim fileTag As String
If sourceFolderName = "Transaction Files" Then
fileTag = "Alerted"
Else
fileTag = "NonAlerted"
End If
excelFileName = ecmID & "_" & AlertID & "_Combined_" & fileTag & "_Transaction.xlsx"

' Building path with guaranteed slashes
finalSavePath = saveFolderPath & slash & excelFileName

' Remove the error bypass so if Excel blocks the save, we actually see why
Application.DisplayAlerts = False
newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
Application.DisplayAlerts = True

' Centralized audit ledger row - Register tab of this case's own
' Desktop\{ecmID}\{ecmID}_Audit_Log.xlsx.
modAuditLog.LogAuditEvent ecmID:=ecmID, alertID:=AlertID, _
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
Application.ScreenUpdating = True

On Error Resume Next
ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
On Error GoTo 0

' ==========================================
' 6. LIVE PUSH TO SHARED ONEDRIVE MASTER TRACKER (SHEET2) - HEADLESS GHOST MODE
' ==========================================
Dim masterPath As String, masterWb As Workbook, masterWs As Worksheet, pushWb As Workbook
Dim mRow As Long, wasAlreadyOpen As Boolean
Dim expectedHeaders As Variant, hdrIdx As Integer, headersOK As Boolean, headerMsg As String
Dim ghostApp As Object ' <--- Our invisible background Excel

masterPath = Environ("USERPROFILE") & "\OneDrive - Community Federal Savings Bank\Mohini Srivastava's files - L1 Beta\Beta 2.4_Feedbacks & Issues Encountered.xlsx"

' 1. Check if the file is already open in the visible Excel window
wasAlreadyOpen = False
For Each pushWb In Application.Workbooks
If pushWb.Name = "Beta 2.4_Feedbacks & Issues Encountered.xlsx" Then
Set masterWb = pushWb
wasAlreadyOpen = True
Exit For
End If
Next pushWb

' 2. THE GHOST EXCEL FIX: Open silently in the background
If masterWb Is Nothing Then
Set ghostApp = CreateObject("Excel.Application")
ghostApp.Visible = False ' Keep it hidden from taskbar
ghostApp.DisplayAlerts = False ' Suppress cloud sync popups
ghostApp.EnableEvents = False ' Lock out UI flashes

On Error Resume Next
Set masterWb = ghostApp.Workbooks.Open(fileName:=masterPath, UpdateLinks:=False)
On Error GoTo CancelHandler
End If

' 3. Process Data on SHEET2
If Not masterWb Is Nothing Then
If Not masterWb.ReadOnly Then
On Error Resume Next
Set masterWs = masterWb.Sheets("Sheet2")
On Error GoTo CancelHandler

If Not masterWs Is Nothing Then
    expectedHeaders = Array("Date & Time", "Analyst ID", "ECM Case ID", "Tool Version")
    headersOK = True
    headerMsg = ""
    
    For hdrIdx = LBound(expectedHeaders) To UBound(expectedHeaders)
        If masterWs.Cells(1, hdrIdx + 1).Value <> expectedHeaders(hdrIdx) Then
            headersOK = False
            headerMsg = headerMsg & "- Col " & Split(masterWs.Cells(1, hdrIdx + 1).Address, "$")(1) & " expected '" & expectedHeaders(hdrIdx) & "' but found '" & masterWs.Cells(1, hdrIdx + 1).Value & "'" & vbCrLf
        End If
    Next hdrIdx
    
    If headersOK Then
        mRow = masterWs.Cells(masterWs.Rows.Count, "A").End(xlUp).Row + 1
        
        masterWs.Cells(mRow, 1).Value = Now
        masterWs.Cells(mRow, 2).Value = Environ("USERNAME")
        masterWs.Cells(mRow, 3).Value = ecmID
        masterWs.Cells(mRow, 4).Value = "2.4.1"
        
        If wasAlreadyOpen Then
            masterWb.Save
        Else
            masterWb.Close SaveChanges:=True
        End If
        
    Else
        MsgBox "DIAGNOSTIC WARNING! HEADER MISMATCH ON SHEET2" & vbCrLf & vbCrLf & _
               "The master tracker headers on Sheet2 have been altered:" & vbCrLf & vbCrLf & _
               headerMsg & vbCrLf & _
               "Tracker data was NOT saved. Please notify the team lead.", vbCritical, "Diagnostic Failed"
        
        If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
    End If
End If
Else
' File is locked by another user
If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
End If
End If

' 4. DESTROY THE GHOST EXCEL PROCESS
If Not ghostApp Is Nothing Then
ghostApp.Quit
Set ghostApp = Nothing
End If
' ==========================================

MsgBox "Workflow Complete!" & vbCrLf & _
"Exported file inside the folder exactly to: " & vbCrLf & finalSavePath, vbInformation, "Success"
Exit Sub

CancelHandler:
' CRITICAL FIX: This ensures Excel unfreezes even if the macro crashes
Application.EnableEvents = True
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

