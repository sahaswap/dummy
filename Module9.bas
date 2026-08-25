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
Dim LastRowSource As Long, LastRowMaster As Long, lastCol As Long
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
ecmID = Trim(wsHome.Range("J9").Value)
AlertID = Trim(wsHome.Range("J10").Value)
If AlertID = "" Then AlertID = "ALERT"

If ecmID = "" Then
MsgBox "Action Denied: ECM ID is missing in J9.", vbCritical, "Missing ID"
Exit Sub
End If

' ==========================================
' 1b. EXPORT FORMAT PICKER - frmExportMode (Legacy / EN Network / Pivot
'     Analysis / Cancel), same pattern as frmSearchMode for OSDD Search.
'     EN Network no longer asks Alerted/Non-Alerted vs Lookback
'     separately - choosing it generates BOTH files in one go (see
'     section 4-EN below).
' ==========================================
Dim exportMode As String
frmExportMode.Show vbModal

If frmExportMode.userCancelled Then
    Unload frmExportMode
    Exit Sub
End If
exportMode = frmExportMode.SelectedMode   ' "LEGACY" / "EN" / "PIVOT"
Unload frmExportMode

' Source folder is "Transaction Files" for every mode except Pivot
' Analysis, which reads from its own separate "\Pivot" folder instead.
Dim sourceFolderName As String
If exportMode = "PIVOT" Then
    sourceFolderName = "Pivot"
Else
    sourceFolderName = "Transaction Files"
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

' ==========================================
' 1c. PIVOT ANALYSIS EXPORT (own source folder: \Pivot)
'   Runs on a PRIVATE scratch sheet of its own - ConsolidatedData is
'   NEVER touched, not even temporarily, unlike every other mode below
'   which stages through it. Combine + cleanup mirrors steps 2-3
'   exactly, just targeting that scratch sheet instead of WsMaster.
'   Output (data + the same 4 pivots) is saved INSIDE \Pivot itself,
'   not the main case folder, named "..._Pivot Analysis.xlsx".
' ==========================================
If exportMode = "PIVOT" Then
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    origCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsPivScratch As Worksheet
    Dim pFile As Object, pWb As Workbook, pWs As Worksheet
    Dim pHeaderCell As Range, pHeaderRow As Long, pHeadCol As Long
    Dim pLastRowSource As Long, pLastRowMaster As Long, pHeaderCopied As Boolean

    ' its own scratch sheet - a completely separate area from ConsolidatedData
    On Error Resume Next
    ThisWorkbook.Sheets("TempPivotScratch").Delete
    On Error GoTo CancelHandler
    Set wsPivScratch = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsPivScratch.Name = "TempPivotScratch"

    ' combine every .xls* in \Pivot, same "Transaction ID" anchor logic as step 2
    pHeaderCopied = False
    For Each pFile In objFolder.Files
        If (InStr(1, pFile.Name, ".xls", vbTextCompare) > 0) And (Left(pFile.Name, 2) <> "~$") And (pFile.Name <> ThisWorkbook.Name) Then
            Set pWb = Workbooks.Open(pFile.path, ReadOnly:=True, UpdateLinks:=False)
            On Error Resume Next
            Set pWs = pWb.Sheets(1)
            On Error GoTo CancelHandler
            If Not pWs Is Nothing Then
                With pWs
                    Set pHeaderCell = .Cells.Find(What:="Transaction ID", LookIn:=xlValues, LookAt:=xlWhole)
                    If Not pHeaderCell Is Nothing Then
                        pHeaderRow = pHeaderCell.row
                        pHeadCol = pHeaderCell.Column
                        pLastRowSource = .Cells(.Rows.count, pHeadCol).End(xlUp).row
                        If pLastRowSource >= pHeaderRow Then
                            If Not pHeaderCopied Then
                                .Range(.Cells(pHeaderRow, pHeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=wsPivScratch.Range("A1")
                                pHeaderCopied = True
                            Else
                                If pLastRowSource > pHeaderRow Then
                                    pLastRowMaster = wsPivScratch.Cells(wsPivScratch.Rows.count, "A").End(xlUp).row + 1
                                    .Range(.Cells(pHeaderRow + 1, pHeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=wsPivScratch.Range("A" & pLastRowMaster)
                                End If
                            End If
                        End If
                    End If
                End With
            End If
            pWb.Close SaveChanges:=False
        End If
    Next pFile

    ' same cleanup the other modes get (real dates, currency format, Counterparty column)
    CleanTransactionData wsPivScratch

    ' output workbook: the data + the same 4 "Legacy" pivots (shared helper)
    Dim newWbPiv As Workbook
    Set newWbPiv = Workbooks.Add
    wsPivScratch.Copy Before:=newWbPiv.Sheets(1): ActiveSheet.Name = "Pivot Data"
    BuildEnPivots newWbPiv, "Pivot Data", "Pivot", "Pivot Data"

    Dim pSh As Long
    For pSh = newWbPiv.Sheets.count To 1 Step -1
        Select Case newWbPiv.Sheets(pSh).Name
            Case "Pivot Data", "Pivot"
                ' keep
            Case Else
                newWbPiv.Sheets(pSh).Delete
        End Select
    Next pSh

    ' remove the scratch sheet - ConsolidatedData was never touched by this mode
    On Error Resume Next
    ThisWorkbook.Sheets("TempPivotScratch").Delete
    On Error GoTo CancelHandler

    ' save INSIDE \Pivot itself (folderPath), not the main case folder
    Dim pivotFileName As String, pivotSavePath As String
    pivotFileName = ecmID & "_" & AlertID & "_Pivot Analysis.xlsx"
    pivotSavePath = folderPath & slash & pivotFileName
    newWbPiv.SaveAs fileName:=pivotSavePath, FileFormat:=51

    newWbPiv.Sheets("Pivot Data").Activate

    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = origCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
    ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
    ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
    Application.OnTime Now + TimeSerial(0, 0, 1), "PushTrxTracker_Deferred"
    On Error GoTo 0

    MsgBox "Pivot Analysis export complete!" & vbCrLf & _
        "ConsolidatedData was not touched." & vbCrLf & _
        "Saved to:" & vbCrLf & pivotSavePath, vbInformation, "Success"
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

' The raw snapshot feeds the "Raw Transactions" sheet. PIVOT mode never
' reaches this line (it exits earlier, before ConsolidatedData is even
' touched). Both remaining modes need it: LEGACY builds Raw Transactions
' directly, and EN Network's combined export always includes the
' Alerted/Non-Alerted file, which also needs it - so this now runs
' unconditionally rather than checking for a "Lookback-only" mode that
' no longer exists (EN Network always builds both files in one go).
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
lastCol = WsMaster.Cells(1, WsMaster.Columns.count).End(xlToLeft).Column + 1

If LastRowMaster > 1 Then
WsMaster.Cells(1, lastCol).Value = "Counterparty"
WsMaster.Range(WsMaster.Cells(2, lastCol), WsMaster.Cells(LastRowMaster, lastCol)).FormulaR1C1 = _
"=IF(RC" & drCrCol & "=""DR"", RC" & benNameCol & ", RC" & origNameCol & ")"

WsMaster.Cells(1, lastCol - 1).Copy
WsMaster.Cells(1, lastCol).PasteSpecial Paste:=xlPasteFormats
WsMaster.Range(WsMaster.Cells(2, lastCol - 1), WsMaster.Cells(LastRowMaster, lastCol - 1)).Copy
WsMaster.Range(WsMaster.Cells(2, lastCol), WsMaster.Cells(LastRowMaster, lastCol)).PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False
End If
End If

' ==========================================
' 4-EN. EN NETWORK EXPORT - generates BOTH files in one go
'   Runs INSTEAD of the Legacy dedupe path when the analyst chose EN
'   Network. There is no longer a sub-picker for this - clicking EN
'   Network builds the Lookback Transactions file FIRST (it needs the
'   FULL Yes+No dataset), then the Alerted/Non-Alerted file SECOND
'   (its own final step is what narrows ConsolidatedData down to
'   Yes-only for good - doing that any earlier would strip the "No"
'   rows Lookback still needs). One combined Excel-state restore and
'   one combined success message cover both files; the Legacy block
'   below is left untouched.
' ==========================================
If exportMode = "EN" Then
    ' ---------------------------------------------------------------
    ' 4-LB. Lookback Transactions (built FIRST - needs full Yes+No)
    '   All transactions (Yes AND No) in the 1-year lookback window:
    '   start = 1st of month, one year back from the LAST alerted date;
    '   end = the last alerted date. Data rows + the same Legacy pivots.
    ' ---------------------------------------------------------------
    Dim aColLB As Long, aDateColLB As Long, scanLastLB As Long, rLB As Long
    Dim keepLB As Boolean, cvLB As Variant
    Dim lastAlertedLB As Date, haveAlertedLB As Boolean, lbStart As Date, lbEnd As Date
    Dim wsLB As Worksheet, lbLastRow As Long
    Dim lbSavedPath As String

    aColLB = 0: aDateColLB = 0
    On Error Resume Next
    aColLB = WsMaster.Rows(1).Find(What:="Is Alerted Transaction?", LookAt:=xlWhole).Column
    aDateColLB = WsMaster.Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
    On Error GoTo CancelHandler
    If aColLB = 0 Then
        MsgBox "EN Network export needs an 'Is Alerted Transaction?' column (exact name), but it wasn't found in the data.", vbCritical, "Column Not Found"
        GoTo CancelHandler
    End If

    ' last alerted date -> window [1st-of-month one year back .. last alerted
    ' date]. Bulk read + LastDataRow, same mechanism the EN block uses below.
    Dim sArrLB As Variant, dArrLB As Variant
    scanLastLB = LastDataRow(WsMaster)
    haveAlertedLB = False
    If scanLastLB > 1 Then
        sArrLB = WsMaster.Range(WsMaster.Cells(2, aColLB), WsMaster.Cells(scanLastLB, aColLB)).Value
        If aDateColLB > 0 Then _
            dArrLB = WsMaster.Range(WsMaster.Cells(2, aDateColLB), WsMaster.Cells(scanLastLB, aDateColLB)).Value
        For rLB = 1 To scanLastLB - 1
            If Trim(CStr(ArrCell(sArrLB, rLB))) = "Yes" Then
                cvLB = ArrCell(dArrLB, rLB)
                If IsDate(cvLB) Then
                    If Not haveAlertedLB Then
                        lastAlertedLB = CDate(cvLB): haveAlertedLB = True
                    ElseIf CDate(cvLB) > lastAlertedLB Then
                        lastAlertedLB = CDate(cvLB)
                    End If
                End If
            End If
        Next rLB
    End If
    If Not haveAlertedLB Then
        MsgBox "No dated 'Yes' alerted transactions were found, so the EN Network export can't be built.", vbCritical, "No Alerted Rows"
        GoTo CancelHandler
    End If
    lbStart = DateSerial(Year(lastAlertedLB) - 1, Month(lastAlertedLB), 1)  ' 1st of month, 1yr back
    lbEnd = lastAlertedLB

    ' build workbook: Lookback Transactions (all Yes+No in window) + pivots
    Set newWb = Workbooks.Add
    WsMaster.Copy Before:=newWb.Sheets(1): ActiveSheet.Name = "Lookback Transactions"
    Set wsLB = newWb.Sheets("Lookback Transactions")
    ' date-window only - keeps BOTH Yes and No rows (splitCol = 0). WsMaster
    ' itself is only ever COPIED FROM here, never modified - it must stay
    ' fully intact (Yes+No) for the Alerted/Non-Alerted build right after.
    FilterRowsFast wsLB, 0, "", aDateColLB, True, lbStart, lbEnd
    With wsLB.Cells
        .WrapText = False: .EntireColumn.AutoFit: .WrapText = True
        .EntireRow.AutoFit: .VerticalAlignment = xlTop
    End With

    BuildEnPivots newWb, "Lookback Transactions", "Pivot", "Lookback Transactions"

    ' drop the default blank sheet(s)
    Application.DisplayAlerts = False
    For rLB = newWb.Sheets.count To 1 Step -1
        Select Case newWb.Sheets(rLB).Name
            Case "Lookback Transactions", "Pivot"
                ' keep
            Case Else
                newWb.Sheets(rLB).Delete
        End Select
    Next rLB
    Application.DisplayAlerts = True

    ' save: {ECM}_{AlertID}_Lookback Transactions (mm.dd.yyyy to mm.dd.yyyy).xlsx
    excelFileName = ecmID & "_" & AlertID & "_Lookback Transactions (" & _
        Format$(lbStart, "mm.dd.yyyy") & " to " & Format$(lbEnd, "mm.dd.yyyy") & ").xlsx"
    finalSavePath = saveFolderPath & slash & excelFileName
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
    Application.DisplayAlerts = True
    lbSavedPath = finalSavePath   ' remember before the Alerted/Non-Alerted save overwrites finalSavePath

    ' NOTE: ConsolidatedData is NOT narrowed to Yes-only here (that used to
    ' happen at this point). It must stay full Yes+No until the
    ' Alerted/Non-Alerted build below has copied from it - THAT block's own
    ' final step is what narrows it, once, for the whole combined export.

    ' ---------------------------------------------------------------
    ' 4-AN. Alerted / Non-Alerted Trx File (built SECOND)
    ' ---------------------------------------------------------------
    Dim wsAlertedEN As Worksheet, wsNonEN As Worksheet, wsPivotEN As Worksheet
    Dim aColEN As Long, aDateColEN As Long, aLastEN As Long, naLastEN As Long
    Dim rEN As Long, scanLastEN As Long, keepEN As Boolean, wsTidy As Variant, cvEN As Variant
    Dim minAlerted As Date, maxAlerted As Date, haveAlerted As Boolean
    Dim winStartEN As Date, winEndEN As Date
    Dim lastRowA As Long, lastColA As Long, ddLastEN As Long
    Dim ptRangeA As Range, ptCacheA As PivotCache, ptRangeT As Range, ptCacheT As PivotCache
    Dim pt3EN As PivotTable, pt4EN As PivotTable

    ' locate the split column + date column on the cleaned master
    aColEN = 0: aDateColEN = 0
    On Error Resume Next
    aColEN = WsMaster.Rows(1).Find(What:="Is Alerted Transaction?", LookAt:=xlWhole).Column
    aDateColEN = WsMaster.Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
    On Error GoTo CancelHandler
    If aColEN = 0 Then
        MsgBox "EN Network export needs an 'Is Alerted Transaction?' column (exact name), but it wasn't found in the data.", vbCritical, "Column Not Found"
        GoTo CancelHandler
    End If

    ' --- alerted date window: min/max Transaction Date of the "Yes" rows,
    '     snapped to first-day-of-month .. last-day-of-month (leap-safe).
    '     The split + date columns are read in ONE bulk call each and scanned
    '     in memory - cell-by-cell COM reads were slow on a full account
    '     history. Last row comes from LastDataRow (all columns), so trailing
    '     rows with a blank split value are no longer missed. ---
    Dim sArrEN As Variant, dArrEN As Variant
    scanLastEN = LastDataRow(WsMaster)
    haveAlerted = False
    If scanLastEN > 1 Then
        sArrEN = WsMaster.Range(WsMaster.Cells(2, aColEN), WsMaster.Cells(scanLastEN, aColEN)).Value
        If aDateColEN > 0 Then _
            dArrEN = WsMaster.Range(WsMaster.Cells(2, aDateColEN), WsMaster.Cells(scanLastEN, aDateColEN)).Value
        For rEN = 1 To scanLastEN - 1
            If Trim(CStr(ArrCell(sArrEN, rEN))) = "Yes" Then
                cvEN = ArrCell(dArrEN, rEN)
                If IsDate(cvEN) Then
                    If Not haveAlerted Then
                        minAlerted = CDate(cvEN): maxAlerted = CDate(cvEN): haveAlerted = True
                    Else
                        If CDate(cvEN) < minAlerted Then minAlerted = CDate(cvEN)
                        If CDate(cvEN) > maxAlerted Then maxAlerted = CDate(cvEN)
                    End If
                End If
            End If
        Next rEN
    End If
    If haveAlerted Then
        winStartEN = DateSerial(Year(minAlerted), Month(minAlerted), 1)
        winEndEN = DateSerial(Year(maxAlerted), Month(maxAlerted) + 1, 0)
    Else
        MsgBox "No 'Yes' alerted transactions were found, so the Non Alerted window can't be built. The Non Alerted sheet will be empty.", vbExclamation, "No Alerted Rows"
    End If

    ' --- build the export workbook ---
    Set newWb = Workbooks.Add
    WsRawTemp.Copy Before:=newWb.Sheets(1): ActiveSheet.Name = "Raw Transactions"
    ' Raw Transactions is a PRE-cleanup snapshot (taken before step 3 ran on
    ' WsMaster), so unlike Alerted/Non-Alerted it never got date conversion,
    ' amount formatting, or the Counterparty column. Run that same cleanup on
    ' it now so it matches the rest instead of showing raw source formatting.
    CleanTransactionData newWb.Sheets("Raw Transactions")

    ' Alerted Transaction = rows where split = "Yes"
    WsMaster.Copy After:=newWb.Sheets(newWb.Sheets.count): ActiveSheet.Name = "Alerted Transaction"
    Set wsAlertedEN = newWb.Sheets("Alerted Transaction")
    FilterRowsFast wsAlertedEN, aColEN, "Yes", 0, False, 0, 0

    ' Non Alerted Transaction = "No" rows whose date is in the alerted-month
    ' window. If there are NO such "No" rows at all (e.g. a single-day alert
    ' with no surrounding non-alerted activity), the sheet is NOT filled with
    ' the "Yes" rows anymore - instead it gets a single message in A1 stating
    ' that 0 non-alerted transactions were found, with the window's dates.
    Dim hasNoInWin As Boolean
    hasNoInWin = False
    If haveAlerted And scanLastEN > 1 Then
        For rEN = 1 To scanLastEN - 1
            If Trim(CStr(ArrCell(sArrEN, rEN))) = "No" Then
                cvEN = ArrCell(dArrEN, rEN)
                If IsDate(cvEN) Then
                    If CDate(cvEN) >= winStartEN And CDate(cvEN) <= winEndEN Then hasNoInWin = True: Exit For
                End If
            End If
        Next rEN
    End If

    WsMaster.Copy After:=newWb.Sheets(newWb.Sheets.count): ActiveSheet.Name = "Non Alerted Transaction"
    Set wsNonEN = newWb.Sheets("Non Alerted Transaction")

    If Not haveAlerted Then
        ' No dated "Yes" rows at all - no window could even be determined
        ' (the earlier MsgBox already flagged this to the analyst).
        wsNonEN.Cells.Clear
        wsNonEN.Range("A1").Value = "No alerted transactions were found, so the Non Alerted window could not be determined."
    ElseIf Not hasNoInWin Then
        ' A real window exists, but zero "No" rows fall inside it.
        wsNonEN.Cells.Clear
        wsNonEN.Range("A1").Value = "There were 0 non-alerted transactions during the alerted month(s) " & _
            Format$(winStartEN, "mm/dd/yyyy") & " to " & Format$(winEndEN, "mm/dd/yyyy") & "."
    Else
        FilterRowsFast wsNonEN, aColEN, "No", aDateColEN, True, winStartEN, winEndEN
    End If

    ' Tidy all three data sheets, and force the SAME "Transaction Date" format
    ' on each. Raw Transactions just got it via CleanTransactionData above;
    ' Alerted/Non-Alerted already have real date VALUES (from WsMaster's step
    ' 3 cleanup) but were showing the long "dddd, mmmm d, yyyy" format there -
    ' this is what made Alerted look different from Non-Alerted. Setting it
    ' explicitly here (rather than relying on BuildEnPivots' side effect,
    ' which only ever touched Alerted Transaction) unifies all three.
    Dim dateColTidy As Long, lastRTidy As Long
    For Each wsTidy In Array("Raw Transactions", "Alerted Transaction", "Non Alerted Transaction")
        With newWb.Sheets(CStr(wsTidy)).Cells
            .WrapText = False: .EntireColumn.AutoFit: .WrapText = True
            .EntireRow.AutoFit: .VerticalAlignment = xlTop
        End With
        On Error Resume Next
        dateColTidy = 0
        dateColTidy = newWb.Sheets(CStr(wsTidy)).Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
        If dateColTidy > 0 Then
            lastRTidy = newWb.Sheets(CStr(wsTidy)).Cells(newWb.Sheets(CStr(wsTidy)).Rows.count, dateColTidy).End(xlUp).row
            If lastRTidy > 1 Then
                newWb.Sheets(CStr(wsTidy)).Range(newWb.Sheets(CStr(wsTidy)).Cells(2, dateColTidy), _
                    newWb.Sheets(CStr(wsTidy)).Cells(lastRTidy, dateColTidy)).NumberFormat = "m/d/yyyy"
            End If
        End If
        On Error GoTo CancelHandler
    Next wsTidy

    ' ---- Pivots (same as Legacy), built from Alerted Transaction ----
    BuildEnPivots newWb, "Alerted Transaction", "Alerted Transaction Pivot", "Alerted Transaction"

    ' drop the default blank sheet(s) that Workbooks.Add created - keep only ours
    Application.DisplayAlerts = False
    For rEN = newWb.Sheets.count To 1 Step -1
        Select Case newWb.Sheets(rEN).Name
            Case "Raw Transactions", "Alerted Transaction", "Alerted Transaction Pivot", "Non Alerted Transaction"
                ' keep
            Case Else
                newWb.Sheets(rEN).Delete
        End Select
    Next rEN
    Application.DisplayAlerts = True

    ' Final sheet order: Raw -> Alerted -> Alerted Pivot -> Non Alerted
    newWb.Sheets("Non Alerted Transaction").Move After:=newWb.Sheets(newWb.Sheets.count)

    ' clean up the raw-backup helper sheet in THIS workbook
    On Error Resume Next
    Application.DisplayAlerts = False
    wsHome.Parent.Sheets("TempRawBackup").Delete
    Application.DisplayAlerts = True
    On Error GoTo CancelHandler

    ' ---- save (same file name as Legacy) ----
    excelFileName = ecmID & "_" & AlertID & "_Combined Alerted & Non Alerted Transactions.xlsx"
    finalSavePath = saveFolderPath & slash & excelFileName
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
    Application.DisplayAlerts = True

    ' ConsolidatedData must hold ONLY the alerted transaction data - never
    ' the Non-Alerted or Lookback rows. This is the ONE point in the whole
    ' combined EN Network export where ConsolidatedData finally gets
    ' narrowed - both files above already finished reading from it.
    WsMaster.Cells.Clear
    newWb.Sheets("Alerted Transaction").UsedRange.Copy Destination:=WsMaster.Range("A1")

    ' Refresh Sheet7's [Rule Name] now too, right as ConsolidatedData gets
    ' its final alerted-only content - so it's already correct if the
    ' analyst looks at Sheet7 before ever running Generate Narrative.
    On Error Resume Next
    Module3.RefreshRuleNameTag
    On Error GoTo CancelHandler

    newWb.Sheets("Raw Transactions").Activate

    ' ---- ONE combined finalize / restore Excel + re-protect for BOTH files ----
    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = origCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
    ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
    ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
    Application.OnTime Now + TimeSerial(0, 0, 1), "PushTrxTracker_Deferred"
    On Error GoTo 0

    MsgBox "EN Network export complete! Both files were generated:" & vbCrLf & vbCrLf & _
        "Lookback Transactions:" & vbCrLf & lbSavedPath & vbCrLf & vbCrLf & _
        "Alerted / Non-Alerted Transactions:" & vbCrLf & finalSavePath, vbInformation, "Success"
    Exit Sub
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

' Refresh Sheet7's [Rule Name] now too, right as ConsolidatedData gets its
' final deduped content - so it's already correct if the analyst looks at
' Sheet7 before ever running Generate Narrative.
On Error Resume Next
Module3.RefreshRuleNameTag
On Error GoTo CancelHandler

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





' ==========================================================
' BuildEnPivots - builds the 4 "Legacy" pivots on a fresh pivot sheet,
' sourced from the given data sheet. Shared by the EN Network (Alerted)
' and Lookback exports so the pivot logic lives in ONE place. It finds
' its own "Transaction Date" column and cleans blank-date rows on the
' source (needed so date grouping doesn't crash), exactly like Legacy.
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
    Set cacheScn = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngScn)
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

    ' Pivots 2-4 (date-grouped) need blank Transaction-Date rows removed so
    ' grouping doesn't crash - but that cleanup must NEVER run against
    ' wsData directly. wsData IS the real deliverable sheet the analyst
    ' opens (e.g. "Alerted Transaction") - deleting rows from it here was
    ' the actual root cause of that sheet coming out completely empty
    ' (and of Pivots 2-4 silently never getting built, since the row-count
    ' this block computed afterward gates whether they run at all). Do the
    ' cleanup on a disposable COPY instead, feed Pivots 2-4 from THAT
    ' copy's cache, then delete the copy - wsData itself is never mutated,
    ' no matter how aggressive or fragile the blank-detection turns out to
    ' be on a thin dataset.
    On Error Resume Next
    wsData.Copy After:=wb.Sheets(wb.Sheets.count)
    Set wsPvScratch = wb.Sheets(wb.Sheets.count)
    wsPvScratch.Visible = xlSheetVeryHidden
    On Error GoTo 0
    If wsPvScratch Is Nothing Then GoTo SkipDateGroupedPivots

    On Error Resume Next
    If dCol > 0 Then
        ddLast = wsPvScratch.Cells(wsPvScratch.Rows.count, dCol).End(xlUp).row
        If ddLast > 1 Then
            wsPvScratch.Range(wsPvScratch.Cells(2, dCol), wsPvScratch.Cells(ddLast, dCol)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            wsPvScratch.Range(wsPvScratch.Cells(2, dCol), wsPvScratch.Cells(ddLast, dCol)).NumberFormat = "m/d/yyyy"
        End If
    End If
    On Error GoTo 0

    lastRow = wsPvScratch.Cells(wsPvScratch.Rows.count, "A").End(xlUp).row
    lastCol = wsPvScratch.Cells(1, wsPvScratch.Columns.count).End(xlToLeft).Column
    If lastRow > 1 Then
        Set rngTmp = wsPvScratch.Range(wsPvScratch.Cells(1, 1), wsPvScratch.Cells(lastRow, lastCol))
        Set cacheTmp = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngTmp)

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

        ' PIVOT 3: ENHANCED TEMPORAL (cloned from PT2)
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

        ' PIVOT 4: DR/CR INVERTED (cloned from PT2)
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
    ' Scratch copy's job is done - remove it. wsData (the real deliverable
    ' sheet) was never touched by any of the above, regardless of whether
    ' Pivots 2-4 built successfully.
    On Error Resume Next
    If Not wsPvScratch Is Nothing Then
        Application.DisplayAlerts = False
        wsPvScratch.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    wsPv.Columns("A:W").AutoFit
End Sub

' ==========================================================
' LastDataRow - the true last row containing anything, across ALL
' columns. The filters used to derive the last row from .End(xlUp) on
' the "Is Alerted Transaction?" column alone, so any trailing row whose
' value in THAT column was blank was never examined and survived the
' filter (a non-"Yes" row could leak into the Alerted sheet).
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

' Reads element idx from a bulk .Value read. A single-cell range returns a
' scalar rather than a 2D array, and an unrequested column is Empty - both
' are handled here so callers can index uniformly.
Private Function ArrCell(ByVal v As Variant, ByVal idx As Long) As Variant
    If IsArray(v) Then
        ArrCell = v(idx, 1)
    Else
        ArrCell = v
    End If
End Function

' ==========================================================
' FilterRowsFast - keeps only the rows that match, deleting the rest in a
' SINGLE operation.
'
' Replaces the old "For r = last To 2 Step -1 : Rows(r).Delete" loops. Those
' cost one COM call + a full row-shift PER ROW, which on a whole-account
' history (tens of thousands of rows) took minutes and looked like a hang.
' This writes a temporary flag column, AutoFilters it, and deletes every
' unwanted row at once.
'
'   splitCol  - "Is Alerted Transaction?" column (0 = don't test it)
'   wantVal   - required value in splitCol ("Yes"/"No"); "" = don't test
'   dateCol   - "Transaction Date" column (0 = don't test it)
'   useWindow - True to also require winStart <= date <= winEnd
' ==========================================================
Private Sub FilterRowsFast(ByVal ws As Worksheet, ByVal splitCol As Long, _
    ByVal wantVal As String, ByVal dateCol As Long, ByVal useWindow As Boolean, _
    ByVal winStart As Date, ByVal winEnd As Date)

    Dim lastRow As Long, lastCol As Long, helperCol As Long, r As Long
    Dim sVals As Variant, dVals As Variant, dv As Variant
    Dim keep As Boolean, anyDelete As Boolean
    Dim flags() As Variant
    Dim delRange As Range

    lastRow = LastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    lastCol = 1
    On Error Resume Next
    lastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    helperCol = lastCol + 1

    ' one bulk read per tested column, then decide every row in memory
    If splitCol > 0 Then sVals = ws.Range(ws.Cells(2, splitCol), ws.Cells(lastRow, splitCol)).Value
    If dateCol > 0 Then dVals = ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)).Value

    ReDim flags(1 To lastRow - 1, 1 To 1)
    anyDelete = False
    For r = 1 To lastRow - 1
        keep = True
        If splitCol > 0 And Len(wantVal) > 0 Then
            If Trim(CStr(ArrCell(sVals, r))) <> wantVal Then keep = False
        End If
        If keep And useWindow Then
            dv = ArrCell(dVals, r)
            If IsDate(dv) Then
                If CDate(dv) < winStart Or CDate(dv) > winEnd Then keep = False
            Else
                keep = False
            End If
        End If
        If keep Then
            flags(r, 1) = "K"
        Else
            flags(r, 1) = "D"
            anyDelete = True
        End If
    Next r

    If Not anyDelete Then Exit Sub    ' nothing to remove - leave the sheet alone

    ' flag column -> filter to "D" -> delete those rows in one shot
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    ws.Cells(1, helperCol).Value = "_flag"
    ws.Range(ws.Cells(2, helperCol), ws.Cells(lastRow, helperCol)).Value = flags
    ws.Range(ws.Cells(1, helperCol), ws.Cells(lastRow, helperCol)).AutoFilter Field:=1, Criteria1:="D"

    On Error Resume Next
    Set delRange = ws.Range(ws.Cells(2, helperCol), ws.Cells(lastRow, helperCol)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not delRange Is Nothing Then delRange.EntireRow.Delete

    On Error Resume Next
    ws.AutoFilterMode = False
    On Error GoTo 0
    ws.Columns(helperCol).Delete
End Sub

' ==========================================================
' CleanTransactionData - runs the SAME cleanup that step 3 (AGGRESSIVE DATA
' CLEANUP) applies to WsMaster, but on any given worksheet. Used to bring
' the "Raw Transactions" sheet (an EN Network snapshot taken BEFORE step 3
' runs) up to the same cleaned state as the sheets copied from WsMaster
' afterwards: real Date values in the Transaction Date column (via
' TextToColumns, since source dates can arrive as text), Transaction Amount
' as currency, and the derived Counterparty column.
'
' Uses "m/d/yyyy" (short form) rather than WsMaster's own "dddd, mmmm d,
' yyyy" - this is the format the pivot-grouping code already relies on
' elsewhere, and standardising on it here is what keeps Raw Transactions /
' Alerted Transaction / Non Alerted Transaction all showing the same date
' type instead of three different ones.
' ==========================================================
Private Sub CleanTransactionData(ByVal ws As Worksheet)
    Dim dH As Range, dLast As Long, dRng As Range
    Dim aH As Range, aLast As Long, aRng As Range
    Dim drC As Range, benC As Range, orgC As Range
    Dim drCol As Long, benCol As Long, orgCol As Long
    Dim lastR As Long, lastC As Long

    On Error Resume Next

    ' Transaction Date -> real date values, unified short format
    Set dH = ws.Rows(1).Find(What:="Transaction Date", LookIn:=xlValues, LookAt:=xlPart)
    If Not dH Is Nothing Then
        dLast = ws.Cells(ws.Rows.count, dH.Column).End(xlUp).row
        If dLast > 1 Then
            Set dRng = ws.Range(ws.Cells(2, dH.Column), ws.Cells(dLast, dH.Column))
            dRng.TextToColumns Destination:=dRng.Cells(1, 1), DataType:=xlDelimited, FieldInfo:=Array(Array(1, 3))
            dRng.NumberFormat = "m/d/yyyy"
        End If
    End If

    ' Transaction Amount -> refreshed values + currency format
    Set aH = ws.Rows(1).Find(What:="Transaction Amount", LookIn:=xlValues, LookAt:=xlPart)
    If Not aH Is Nothing Then
        aLast = ws.Cells(ws.Rows.count, aH.Column).End(xlUp).row
        If aLast > 1 Then
            Set aRng = ws.Range(ws.Cells(2, aH.Column), ws.Cells(aLast, aH.Column))
            aRng.Value = aRng.Value
            aRng.NumberFormat = "$#,##0.00"
        End If
    End If

    ' Counterparty = IF(Dr Cr = "DR", Beneficiary Name, Originator Name)
    Set drC = ws.Rows(1).Find(What:="Dr Cr", LookAt:=xlPart)
    Set benC = ws.Rows(1).Find(What:="Beneficiary Name", LookAt:=xlPart)
    Set orgC = ws.Rows(1).Find(What:="Originator Name", LookAt:=xlPart)
    If Not drC Is Nothing And Not benC Is Nothing And Not orgC Is Nothing Then
        drCol = drC.Column: benCol = benC.Column: orgCol = orgC.Column
        lastR = ws.Cells(ws.Rows.count, "A").End(xlUp).row
        lastC = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        If lastR > 1 Then
            ws.Cells(1, lastC).Value = "Counterparty"
            ws.Range(ws.Cells(2, lastC), ws.Cells(lastR, lastC)).FormulaR1C1 = _
                "=IF(RC" & drCol & "=""DR"", RC" & benCol & ", RC" & orgCol & ")"
            ws.Cells(1, lastC - 1).Copy
            ws.Cells(1, lastC).PasteSpecial Paste:=xlPasteFormats
            ws.Range(ws.Cells(2, lastC - 1), ws.Cells(lastR, lastC - 1)).Copy
            ws.Range(ws.Cells(2, lastC), ws.Cells(lastR, lastC)).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        End If
    End If

    On Error GoTo 0
End Sub
