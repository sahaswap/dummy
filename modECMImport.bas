Attribute VB_Name = "modECMImport"
Option Explicit

' ==========================================================================
' IMPORT ECM INFO
'
' Takes the ECM_Alert_Info_<date>_<time>.csv that the "ECM Export" browser
' button saves to Downloads (one row per alert) and, for an alert:
'   1. fills Sheet1:
'        J9  = ECM Case           <- ECM ID
'        J10 = Alert ID           <- Velocity Alert ID
'        J13 = Customer Name      <- Customer Name
'        J14 = Customer Address   <- Address (street, city, state zip, country)
'        L27 = ISO Code 2         <- country code of the address; Sheet1's own
'                                    change handler fills G27 / O27 and the risk
'                                    formulas follow (filled here if it does not)
'      Cells are written one at a time with events on, so Sheet1 applies its
'      usual auto-capitalisation exactly as if they had been typed;
'   2. creates the case folders (same rule as START);
'   3. copies the alert's transaction file from Downloads into
'      <Desktop>\<ECM>\Transaction Files;
'   4. runs "Export Trx File" in AUTO_EXPORT_MODE (EN Network), no picker.
'
' Failsafe - nothing already made for a case is made again:
'   - a Transaction Files folder that already has an Excel file is not copied into;
'   - if the case folder already has its EN "Combined Alerted & Non Alerted
'     Transactions" file, the export is not re-run; instead ConsolidatedData is
'     reloaded from that file's "Alerted Transaction" sheet (what the export
'     itself would put there), so the narrative still matches the case.
'   (The Export Trx File button itself is unchanged and always rebuilds.)
'
' When a queued case is ready an on-top "case ready" message appears above
' every window (NOTIFY_WHEN_READY); clicking OK leaves you in the tool on Sheet1.
'
' Two ways in:
'   AUTOMATIC  The workbook watches Downloads (started on open by Auto_Open)
'              and works as a queue, ONE CASE AT A TIME: whenever the form is
'              empty (after Reset) it loads the next waiting alert - steps 1-4,
'              no prompts - and then leaves it for you to work. Press Reset
'              when you are done and the next alert loads within ~10 seconds.
'              Waiting alerts come from every ECM_Alert_Info csv in Downloads
'              (moved to Downloads\ECM_Imported when picked up), oldest first;
'              an alert counts as done once its Desktop case folder exists.
'              A csv whose alerts are all done is renamed Done_<name>.
'              StopEcmWatcher / StartEcmWatcher pause and resume it.
'   MANUAL     START with an empty ECM ID, or Alt+F8 > Import_ECM_Info:
'              pick the csv and the alert.
' ==========================================================================

Private Const PWD As String = "p7ss"
Private Const NOTIFY_WHEN_READY As Boolean = True ' on-top "case ready" message when a queued case loads
' MsgBox flags: system modal + set foreground + topmost, so the message shows
' above Edge even though Excel is in the background.
Private Const ON_TOP As Long = 4096 + 65536 + 262144    ' vbSystemModal + vbMsgBoxSetForeground + topmost
Private Const MAX_LISTED As Long = 15
Private Const AUTO_EXPORT_MODE As String = "EN"   ' "EN" / "LEGACY" / "PIVOT" / "" = do not export
Private Const WATCH_SECONDS As Long = 10          ' how often Downloads is checked
Private Const IMPORTED_FOLDER As String = "ECM_Imported"
Private Const DONE_PREFIX As String = "Done_"
Private Const LOG_NAME As String = "ECM_import_log.txt"

' Loaded csv
Private mRows As Collection
Private mCEcm As Long, mCVel As Long, mCCust As Long, mCAddr As Long, mCCountry As Long, mCFile As Long

' Watcher state
Private mWatching As Boolean
Private mNextTick As Date
Private mBusy As Boolean

' ==========================================================================
' MANUAL
' ==========================================================================
Public Sub Import_ECM_Info()
    Dim ws As Worksheet, csvPath As String, loadErr As String
    Dim pick As Long, ecmID As String, note As String, copied As String, result As String

    Set ws = ThisWorkbook.Sheets("Sheet1")

    csvPath = ChooseCsv()
    If csvPath = "" Then Exit Sub

    On Error GoTo Fail
    loadErr = LoadCsv(csvPath)
    If loadErr <> "" Then MsgBox loadErr, vbExclamation, "Import ECM Info": Exit Sub

    pick = ChooseAlert(mRows, mCEcm, mCVel, mCCust, Trim$(CStr(ws.Range("J9").Value)))
    If pick < 2 Then Exit Sub
    ecmID = Field(mRows(pick), mCEcm)

    If Trim$(CStr(ws.Range("J9").Value)) <> "" And Trim$(CStr(ws.Range("J9").Value)) <> ecmID Then
        If MsgBox("Sheet1 already holds ECM case " & ws.Range("J9").Value & "." & vbCrLf & vbCrLf & _
                  "Replace the case details with ECM case " & ecmID & "?" & vbCrLf & _
                  "(Tip: press Reset first to clear the rest of the form.)", _
                  vbYesNo + vbQuestion, "Import ECM Info") <> vbYes Then Exit Sub
    End If

    copied = ImportRow(ws, pick, csvPath, note)

    If copied <> "" And AUTO_EXPORT_MODE <> "" Then
        ' Straight on to the export - no prompt. Its own completion message ends the run.
        Application.StatusBar = "ECM case " & ecmID & " imported - running Export Trx File (" & _
                                ExportModeName(AUTO_EXPORT_MODE) & ")..."
        result = RunAutoExport(ecmID, False)
        Application.StatusBar = False
        If result <> "OK" And Left$(result, 2) = "OK" Then
            MsgBox note & vbCrLf & "EN Network files already exist for this case - " & Mid$(result, 6) & ".", _
                   vbInformation, "Import ECM Info"
        ElseIf result <> "OK" And result <> "" Then
            MsgBox note & vbCrLf & result, vbExclamation, "Import ECM Info"
        End If
    Else
        MsgBox note & vbCrLf & "Next: choose Decision and Template, then click Export Trx File.", _
               vbInformation, "Import ECM Info"
    End If
    Exit Sub

Fail:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = True
    On Error GoTo 0
    MsgBox "Import stopped with an error:" & vbCrLf & vbCrLf & "Err " & errNum & ": " & errDesc, _
           vbCritical, "Import ECM Info"
End Sub

' ==========================================================================
' AUTOMATIC - Downloads watcher, one case at a time
' ==========================================================================
Public Sub Auto_Open()
    StartEcmWatcher
End Sub

Public Sub Auto_Close()
    StopEcmWatcher
End Sub

Public Sub StartEcmWatcher()
    mWatching = True
    ScheduleTick
    Application.StatusBar = "ECM auto-import is on - the next waiting alert loads whenever the form is empty."
End Sub

Public Sub StopEcmWatcher()
    mWatching = False
    On Error Resume Next
    Application.OnTime mNextTick, MacroRef("ECM_Watch_Tick"), , False
    Application.StatusBar = False
End Sub

Private Sub ScheduleTick()
    mNextTick = Now + TimeSerial(0, 0, WATCH_SECONDS)
    Application.OnTime mNextTick, MacroRef("ECM_Watch_Tick")
End Sub

' OnTime needs the workbook name when several workbooks are open.
Private Function MacroRef(ByVal procName As String) As String
    MacroRef = "'" & ThisWorkbook.Name & "'!" & procName
End Function

Public Sub ECM_Watch_Tick()
    If Not mWatching Then Exit Sub
    If Not mBusy Then
        mBusy = True
        On Error Resume Next      ' a failed check must never stop the watcher
        CheckQueue
        On Error GoTo 0
        mBusy = False
    End If
    If mWatching Then ScheduleTick
End Sub

' Picks up new exports, then - only if the form is empty - loads the next
' waiting alert. A case on the form (J9 filled) is never replaced.
Private Sub CheckQueue()
    Dim ws As Worksheet, csv As String, idx As Long, waiting As Long
    Set ws = ThisWorkbook.Sheets("Sheet1")

    PickUpNewExports
    waiting = NextWaiting(csv, idx)

    If Trim$(CStr(ws.Range("J9").Value)) <> "" Then
        If waiting > 0 Then
            Application.StatusBar = waiting & " ECM alert(s) waiting - press Reset when you finish case " & _
                                    ws.Range("J9").Value & " to load the next one."
        End If
        Exit Sub
    End If
    If waiting = 0 Then Exit Sub

    LoadQueued ws, csv, idx, waiting - 1
End Sub

' Moves finished ECM_Alert_Info csvs from Downloads into ECM_Imported.
' A file still being written (changed in the last 3 seconds) is left for the next check.
Private Sub PickUpNewExports()
    Dim fso As Object, folder As String, destFolder As String, f As String, dest As String
    Dim found As New Collection, item As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = DownloadsFolder()
    destFolder = folder & "\" & IMPORTED_FOLDER

    f = Dir$(folder & "\ECM_Alert_Info_*.csv")
    Do While f <> ""
        If Now - FileDateTime(folder & "\" & f) > TimeSerial(0, 0, 3) Then found.Add f
        f = Dir$()
    Loop
    If found.Count = 0 Then Exit Sub

    If Not fso.FolderExists(destFolder) Then fso.CreateFolder destFolder
    For Each item In found
        dest = destFolder & "\" & item
        If fso.FileExists(dest) Then dest = destFolder & "\" & fso.GetBaseName(item) & "_" & Format$(Now, "hhnnss") & ".csv"
        fso.MoveFile folder & "\" & item, dest
    Next item
End Sub

' Finds the first waiting alert (oldest csv first, then file order): one whose
' Desktop case folder does not exist yet. Sets csv / idx to it and returns how
' many alerts are waiting in total. Csvs with nothing left are renamed Done_.
Private Function NextWaiting(ByRef csv As String, ByRef idx As Long) As Long
    Dim fso As Object, folder As String, f As String, files As New Collection
    Dim i As Long, j As Long, r As Long, pending As Long, total As Long
    Dim paths() As String, times() As Date, tmpS As String, tmpT As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = DownloadsFolder() & "\" & IMPORTED_FOLDER
    csv = "": idx = 0
    If Not fso.FolderExists(folder) Then Exit Function

    f = Dir$(folder & "\ECM_Alert_Info_*.csv")
    Do While f <> ""
        files.Add folder & "\" & f
        f = Dir$()
    Loop
    If files.Count = 0 Then Exit Function

    ' oldest first
    ReDim paths(1 To files.Count): ReDim times(1 To files.Count)
    For i = 1 To files.Count
        paths(i) = files(i): times(i) = FileDateTime(files(i))
    Next i
    For i = 1 To files.Count - 1
        For j = i + 1 To files.Count
            If times(j) < times(i) Then
                tmpS = paths(i): paths(i) = paths(j): paths(j) = tmpS
                tmpT = times(i): times(i) = times(j): times(j) = tmpT
            End If
        Next j
    Next i

    For i = 1 To files.Count
        If LoadCsv(paths(i)) = "" Then
            pending = 0
            For r = 2 To mRows.Count
                If Not fso.FolderExists(CaseFolderPath(Field(mRows(r), mCEcm))) Then
                    pending = pending + 1
                    If csv = "" Then csv = paths(i): idx = r
                End If
            Next r
            If pending = 0 Then
                fso.MoveFile paths(i), folder & "\" & DONE_PREFIX & fso.GetFileName(paths(i))
            End If
            total = total + pending
        End If
    Next i
    NextWaiting = total
End Function

' Steps 1-4 for one queued alert, with no prompts. Leaves the case on Sheet1
' for the analyst. Only a failure shows a message.
Private Sub LoadQueued(ByVal ws As Worksheet, ByVal csv As String, ByVal idx As Long, ByVal stillWaiting As Long)
    Dim ecmID As String, note As String, copied As String, result As String

    If LoadCsv(csv) <> "" Then Exit Sub
    ecmID = Field(mRows(idx), mCEcm)

    ThisWorkbook.Activate
    ws.Activate
    Application.StatusBar = "ECM auto-import: loading ECM " & ecmID & "..."

    On Error GoTo Failed
    copied = ImportRow(ws, idx, csv, note)
    If copied = "" Then
        result = "transaction file not found in Downloads"
    ElseIf AUTO_EXPORT_MODE = "" Then
        result = "OK"
    Else
        result = RunAutoExport(ecmID, True)
        If result = "" Then result = "export did not finish (see its message)"
    End If
    GoTo Report

Failed:
    result = "error " & Err.Number & ": " & Err.Description
    ProtectSheet ws
    Application.EnableEvents = True
    Resume Report

Report:
    WriteLog ecmID, result
    PrepareToolWindow ws
    If Left$(result, 2) = "OK" Then
        Application.StatusBar = "ECM " & ecmID & IIf(result = "OK", " loaded and exported. ", " loaded (EN files already existed, not recreated). ") & _
            IIf(stillWaiting > 0, stillWaiting & " more waiting - press Reset when done to load the next.", "No more alerts waiting.")
        If NOTIFY_WHEN_READY Then
            MsgBox "ECM case " & ecmID & " is ready in the tool." & vbCrLf & _
                   IIf(stillWaiting > 0, stillWaiting & " more waiting after this one.", "No more alerts waiting.") & vbCrLf & vbCrLf & _
                   "Click OK to work it.", vbInformation Or ON_TOP, "ECM case ready"
        End If
    Else
        Application.StatusBar = False
        MsgBox "ECM auto-import - case " & ecmID & ":" & vbCrLf & result & vbCrLf & vbCrLf & note, _
               vbExclamation Or ON_TOP, "ECM auto-import"
    End If
End Sub

' Makes the tool the active workbook on Sheet1 and un-minimises it, so that
' when the on-top message is closed you land straight in the case.
Private Sub PrepareToolWindow(ByVal ws As Worksheet)
    Const WIN_MINIMIZED As Long = -4140, WIN_NORMAL As Long = -4143   ' xlMinimized, xlNormal
    On Error Resume Next
    ThisWorkbook.Activate
    ThisWorkbook.Windows(1).Activate
    If ThisWorkbook.Windows(1).WindowState = WIN_MINIMIZED Then ThisWorkbook.Windows(1).WindowState = WIN_NORMAL
    If Application.WindowState = WIN_MINIMIZED Then Application.WindowState = WIN_NORMAL
    ws.Activate
End Sub

Private Sub WriteLog(ByVal ecmID As String, ByVal result As String)
    Dim fso As Object, ts As Object, path As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = DownloadsFolder() & "\" & IMPORTED_FOLDER & "\" & LOG_NAME
    Set ts = fso.OpenTextFile(path, 8, True)     ' 8 = append, create if missing
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & ecmID & vbTab & result
    ts.Close
End Sub

' ==========================================================================
' SHARED
' ==========================================================================

' Reads the csv into mRows and finds the columns. Returns "" or an error text.
Private Function LoadCsv(ByVal csvPath As String) As String
    Set mRows = ParseCsv(ReadUtf8(csvPath))
    If mRows.Count < 2 Then LoadCsv = "That file has no alert rows:" & vbCrLf & csvPath: Exit Function
    mCEcm = ColIdx(mRows(1), "ECM ID")
    mCVel = ColIdx(mRows(1), "Velocity Alert ID")
    mCCust = ColIdx(mRows(1), "Customer Name")
    mCAddr = ColIdx(mRows(1), "Address")
    mCCountry = ColIdx(mRows(1), "Country")
    mCFile = ColIdx(mRows(1), "Transaction File")
    If mCEcm < 0 Then LoadCsv = "This does not look like an ECM export file (no 'ECM ID' column):" & vbCrLf & csvPath
End Function

' Steps 1-3 for row idx of mRows: fill Sheet1, create the folders, copy the
' transaction file. Returns the copied file's path ("" if not found); note
' gets a readable summary.
Private Function ImportRow(ByVal ws As Worksheet, ByVal idx As Long, ByVal csvPath As String, _
                           ByRef note As String) As String
    Dim rec As Variant, ecmID As String, velID As String, custName As String
    Dim addr As String, iso2 As String, caseFolder As String, copied As String

    rec = mRows(idx)
    ecmID = Field(rec, mCEcm)
    velID = Field(rec, mCVel)
    custName = Field(rec, mCCust)
    addr = Field(rec, mCAddr)
    iso2 = CountryCode(Field(rec, mCCountry))

    Application.EnableEvents = True
    ' Clear what the previous case put there, so nothing carries over.
    PutValue ws, "J13", ""
    PutValue ws, "J14", ""
    PutValue ws, "L27", ""
    PutValue ws, "G27", ""
    PutValue ws, "O27", ""

    PutValue ws, "J9", ecmID
    PutValue ws, "J10", velID
    If custName <> "" Then PutValue ws, "J13", custName
    If addr <> "" Then PutValue ws, "J14", addr
    If iso2 <> "" Then
        PutValue ws, "L27", iso2
        ' Sheet1's handler normally fills G27/O27 from L27; fill them here if it did not.
        If Trim$(CStr(ws.Range("G27").Value)) = "" Then FillCountryFromIso2 ws, iso2
    End If
    ProtectSheet ws

    note = "Filled Sheet1 for ECM case " & ecmID & ":" & vbCrLf & _
           "   ECM Case (J9), Alert ID (J10)" & _
           IIf(custName <> "", ", Customer Name (J13)", "") & _
           IIf(addr <> "", ", Customer Address (J14)", "") & vbCrLf
    If iso2 <> "" Then
        If Trim$(CStr(ws.Range("G27").Value)) <> "" Then
            note = note & "   Country (G27): " & ws.Range("G27").Value & vbCrLf
        Else
            note = note & "   Country code " & iso2 & " is not in the country table - pick the country in G27." & vbCrLf
        End If
    End If

    caseFolder = MakeCaseFolders(ecmID)
    copied = FirstExcelIn(caseFolder & "\Transaction Files")
    If copied <> "" Then
        note = note & vbCrLf & "Transaction file already in the case folder (not copied again):" & vbCrLf & _
               "   " & copied & vbCrLf
        ImportRow = copied
        Exit Function
    End If
    copied = CopyTransactionFile(csvPath, Field(rec, mCFile), ecmID, caseFolder & "\Transaction Files")
    If copied <> "" Then
        note = note & vbCrLf & "Transaction file copied to:" & vbCrLf & "   " & copied & vbCrLf
    Else
        note = note & vbCrLf & "Transaction file NOT found in Downloads - copy it into:" & vbCrLf & _
               "   " & caseFolder & "\Transaction Files" & vbCrLf
    End If
    ImportRow = copied
End Function

' Step 4. quiet = no completion message and the export's files are closed
' (queue mode). Returns "OK" when the export finished, "OK - <what>" when the
' EN files already existed and were reused, "" or an error text otherwise.
Private Function RunAutoExport(ByVal ecmID As String, ByVal quiet As Boolean) As String
    Dim existing As String
    If UCase$(AUTO_EXPORT_MODE) = "EN" Then
        existing = ExistingEnFile(ecmID)
        If existing <> "" Then
            RunAutoExport = ReuseEnFile(existing)
            Exit Function
        End If
    End If

    gExportLastResult = ""
    gExportQuiet = quiet
    gExportModeOverride = AUTO_EXPORT_MODE
    Consolidated_AML_Workflow
    gExportModeOverride = ""
    gExportQuiet = False
    RunAutoExport = gExportLastResult
End Function

' The case's EN "Combined Alerted & Non Alerted Transactions" file, if one was made before.
Private Function ExistingEnFile(ByVal ecmID As String) As String
    Dim folder As String, f As String
    folder = CaseFolderPath(ecmID)
    f = Dir$(folder & "\" & ecmID & "_*Combined Alerted & Non Alerted Transactions*.xls*")
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then ExistingEnFile = folder & "\" & f: Exit Function
        f = Dir$()
    Loop
End Function

' Instead of rebuilding: fill ConsolidatedData from the existing file's
' "Alerted Transaction" sheet, exactly as the EN export does, and refresh the
' rule-name tag. The file itself is opened read-only and left untouched.
Private Function ReuseEnFile(ByVal path As String) As String
    Dim wb As Workbook, cd As Worksheet, wasOpen As Boolean, prevUpd As Boolean
    prevUpd = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then wasOpen = True: Exit For
    Next wb
    If Not wasOpen Then Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=False)

    Set cd = ThisWorkbook.Sheets("ConsolidatedData")
    ThisWorkbook.Unprotect Password:=PWD
    cd.Unprotect Password:=PWD
    cd.Cells.Clear
    wb.Sheets("Alerted Transaction").UsedRange.Copy Destination:=cd.Range("A1")
    If Not wasOpen Then wb.Close SaveChanges:=False

    On Error Resume Next
    RefreshRuleNameTag
    cd.Protect Password:=PWD
    ThisWorkbook.Protect Password:=PWD, Structure:=True
    ThisWorkbook.Sheets("Sheet1").Activate
    On Error GoTo 0
    Application.ScreenUpdating = prevUpd
    ReuseEnFile = "OK - existing EN file reused, not recreated"
    Exit Function

Fail:
    ReuseEnFile = "could not reuse existing EN file (" & Err.Description & "): " & path
    On Error Resume Next
    If Not wasOpen And Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Protect Password:=PWD, Structure:=True
    Application.ScreenUpdating = prevUpd
End Function

' Closes an open workbook by full path, without saving (it was just saved).
Public Sub CloseSavedExport(ByVal fullPath As String)
    Dim wb As Workbook
    If fullPath = "" Then Exit Sub
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next wb
End Sub

Private Function DownloadsFolder() As String
    DownloadsFolder = Environ$("USERPROFILE") & "\Downloads"
End Function

' --------------------------------------------------------------------------
' Newest ECM_Alert_Info_*.csv in Downloads (with a Yes/No to use it), or a
' file picker. Returns "" if cancelled.
' --------------------------------------------------------------------------
Private Function ChooseCsv() As String
    Dim folder As String, newest As String, imported As String, ans As VbMsgBoxResult
    Dim picked As Variant

    folder = DownloadsFolder()
    newest = NewestCsv(folder)
    imported = NewestCsv(folder & "\" & IMPORTED_FOLDER)
    If imported <> "" Then
        If newest = "" Then
            newest = imported
        ElseIf FileDateTime(imported) > FileDateTime(newest) Then
            newest = imported
        End If
    End If

    If newest <> "" Then
        ans = MsgBox("Use the latest ECM export file?" & vbCrLf & vbCrLf & _
                     Mid$(newest, InStrRev(newest, "\") + 1) & vbCrLf & _
                     "saved " & Format$(FileDateTime(newest), "dd mmm yyyy hh:nn") & vbCrLf & vbCrLf & _
                     "Yes = use it    No = choose another file", _
                     vbYesNoCancel + vbQuestion, "Import ECM Info")
        If ans = vbYes Then ChooseCsv = newest: Exit Function
        If ans = vbCancel Then Exit Function
    End If

    On Error Resume Next
    ChDrive Left$(folder, 1)
    ChDir folder
    On Error GoTo 0
    picked = Application.GetOpenFilename("ECM export (*.csv),*.csv", , "Choose the ECM_Alert_Info file")
    If VarType(picked) = vbBoolean Then Exit Function
    ChooseCsv = CStr(picked)
End Function

Private Function NewestCsv(ByVal folder As String) As String
    Dim f As String, best As String, bestT As Date, t As Date
    On Error Resume Next
    f = Dir$(folder & "\*ECM_Alert_Info_*.csv")       ' includes Done_ files
    Do While f <> ""
        t = FileDateTime(folder & "\" & f)
        If best = "" Or t > bestT Then best = folder & "\" & f: bestT = t
        f = Dir$()
    Loop
    NewestCsv = best
End Function

' --------------------------------------------------------------------------
' Pick one alert row. One row: taken as is. Several: numbered list, the row
' matching the ECM ID already in J9 is the default. Returns the row index in
' rows (2..n), or 0 if cancelled.
' --------------------------------------------------------------------------
Private Function ChooseAlert(ByVal rows As Collection, ByVal cEcm As Long, ByVal cVel As Long, _
                             ByVal cCust As Long, ByVal currentEcm As String) As Long
    Dim i As Long, listText As String, def As String, ans As String

    If rows.Count = 2 Then ChooseAlert = 2: Exit Function

    def = "1"
    For i = 2 To rows.Count
        If currentEcm <> "" And Field(rows(i), cEcm) = currentEcm Then def = CStr(i - 1)
        If i - 1 <= MAX_LISTED Then
            listText = listText & (i - 1) & ")  " & Field(rows(i), cEcm) & "   " & Field(rows(i), cVel) & _
                       "   " & Left$(Field(rows(i), cCust), 28) & vbCrLf
        End If
    Next i
    If rows.Count - 1 > MAX_LISTED Then
        listText = listText & "... and " & (rows.Count - 1 - MAX_LISTED) & " more (type the ECM ID)" & vbCrLf
    End If

    Do
        ans = Trim$(InputBox("Which alert? Type its number or ECM ID." & vbCrLf & vbCrLf & listText, _
                             "Import ECM Info", def))
        If ans = "" Then Exit Function
        If IsNumeric(ans) And Len(ans) <= 3 Then
            If CLng(ans) >= 1 And CLng(ans) <= rows.Count - 1 Then
                ChooseAlert = CLng(ans) + 1
                Exit Function
            End If
        End If
        For i = 2 To rows.Count
            If Field(rows(i), cEcm) = ans Then ChooseAlert = i: Exit Function
        Next i
        MsgBox "No alert '" & ans & "' in this file.", vbExclamation, "Import ECM Info"
    Loop
End Function

' --------------------------------------------------------------------------
' Sheet1 helpers. Sheet1's change handler re-protects the sheet and the
' workbook after every edit, so unprotect before each write.
' --------------------------------------------------------------------------
Private Sub PutValue(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal v As String)
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PWD
    ws.Unprotect Password:=PWD
    On Error GoTo 0
    ws.Range(cellAddr).Value = v
End Sub

Private Sub ProtectSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=PWD
    ThisWorkbook.Protect Password:=PWD, Structure:=True
    On Error GoTo 0
End Sub

' Country table on Sheet1: O37:O290 = name, P = ISO2, Q = ISO3.
Private Sub FillCountryFromIso2(ByVal ws As Worksheet, ByVal iso2 As String)
    Dim m As Variant
    m = Application.Match(iso2, ws.Range("P37:P290"), 0)
    If IsError(m) Then Exit Sub
    PutValue ws, "G27", CStr(ws.Range("O37:O290").Cells(CLng(m), 1).Value)
    PutValue ws, "O27", CStr(ws.Range("Q37:Q290").Cells(CLng(m), 1).Value)
End Sub

Private Function ExportModeName(ByVal mode As String) As String
    Select Case UCase$(mode)
        Case "EN": ExportModeName = "EN Network"
        Case "PIVOT": ExportModeName = "Pivot Analysis"
        Case Else: ExportModeName = "Legacy"
    End Select
End Function

' "US - United States" -> "US"; a bare 2-letter code is returned as is.
Private Function CountryCode(ByVal v As String) As String
    Dim p As Long
    v = Trim$(v)
    p = InStr(v, " - ")
    If p > 0 Then v = Trim$(Left$(v, p - 1))
    If Len(v) = 2 Then CountryCode = UCase$(v)
End Function

' --------------------------------------------------------------------------
' Case folders - same location rule as START (Start_Button_Create_Folders).
' --------------------------------------------------------------------------
Private Function CaseFolderPath(ByVal ecmID As String) As String
    Dim fso As Object, userProfile As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    userProfile = Environ$("USERPROFILE")
    If fso.FolderExists(userProfile & "\OneDrive - Community Federal Savings Bank\Desktop") Then
        CaseFolderPath = userProfile & "\OneDrive - Community Federal Savings Bank\Desktop\" & ecmID
    Else
        CaseFolderPath = userProfile & "\Desktop\" & ecmID
    End If
End Function

Private Function MakeCaseFolders(ByVal ecmID As String) As String
    Dim fso As Object, mainFolder As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    mainFolder = CaseFolderPath(ecmID)
    If Not fso.FolderExists(mainFolder) Then fso.CreateFolder mainFolder
    If Not fso.FolderExists(mainFolder & "\Transaction Files") Then fso.CreateFolder mainFolder & "\Transaction Files"
    If Not fso.FolderExists(mainFolder & "\Pivot") Then fso.CreateFolder mainFolder & "\Pivot"
    MakeCaseFolders = mainFolder
End Function

' First Excel file in a folder (ignores Excel's ~$ lock files), or "".
Private Function FirstExcelIn(ByVal folder As String) As String
    Dim f As String
    f = Dir$(folder & "\*.xls*")
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then FirstExcelIn = folder & "\" & f: Exit Function
        f = Dir$()
    Loop
End Function

' Copies the alert's transaction file (named in the CSV, else the newest
' "<ECM ID>_*.xls*") into destFolder, looking next to the CSV and in
' Downloads. Returns the new path, or "".
Private Function CopyTransactionFile(ByVal csvPath As String, ByVal trxFile As String, _
                                     ByVal ecmID As String, ByVal destFolder As String) As String
    Dim fso As Object, folders(1 To 2) As String, src As String, f As String
    Dim bestT As Date, t As Date, k As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    folders(1) = Left$(csvPath, InStrRev(csvPath, "\") - 1)
    folders(2) = DownloadsFolder()

    For k = 1 To 2
        If trxFile <> "" And src = "" Then
            If fso.FileExists(folders(k) & "\" & trxFile) Then src = folders(k) & "\" & trxFile
        End If
    Next k
    If src = "" Then
        For k = 1 To 2
            f = Dir$(folders(k) & "\" & ecmID & "_*.xls*")
            Do While f <> ""
                t = FileDateTime(folders(k) & "\" & f)
                If src = "" Or t > bestT Then src = folders(k) & "\" & f: bestT = t
                f = Dir$()
            Loop
        Next k
    End If
    If src = "" Then Exit Function

    fso.CopyFile src, destFolder & "\", False     ' never overwrite
    CopyTransactionFile = destFolder & "\" & fso.GetFileName(src)
End Function

' --------------------------------------------------------------------------
' CSV reading
' --------------------------------------------------------------------------
Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object, s As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2            ' text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    s = st.ReadText
    st.Close
    If Left$(s, 1) = ChrW$(65279) Then s = Mid$(s, 2)
    ReadUtf8 = s
End Function

' RFC 4180 CSV: quoted fields, "" inside quotes, commas and line breaks
' inside quotes. Returns a Collection of 0-based string arrays.
Private Function ParseCsv(ByVal txt As String) As Collection
    Dim result As New Collection, fields As Collection
    Dim i As Long, n As Long, ch As String, cur As String, inQ As Boolean

    Set fields = New Collection
    n = Len(txt)
    For i = 1 To n
        ch = Mid$(txt, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(txt, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQ = True
                Case ","
                    fields.Add cur: cur = ""
                Case vbCr
                    ' ignored; vbLf ends the row
                Case vbLf
                    fields.Add cur: cur = ""
                    result.Add ToArray(fields)
                    Set fields = New Collection
                Case Else
                    cur = cur & ch
            End Select
        End If
    Next i
    If cur <> "" Or fields.Count > 0 Then
        fields.Add cur
        result.Add ToArray(fields)
    End If
    Set ParseCsv = result
End Function

Private Function ToArray(ByVal c As Collection) As Variant
    Dim a() As String, i As Long
    ReDim a(0 To c.Count - 1)
    For i = 1 To c.Count
        a(i - 1) = c(i)
    Next i
    ToArray = a
End Function

Private Function ColIdx(ByVal hdr As Variant, ByVal name As String) As Long
    Dim i As Long
    ColIdx = -1
    For i = LBound(hdr) To UBound(hdr)
        If StrComp(Trim$(hdr(i)), name, vbTextCompare) = 0 Then ColIdx = i: Exit Function
    Next i
End Function

' Field i of a row, trimmed; "" when the column is missing. Removes the
' leading apostrophe the export adds to values starting with = + - @.
Private Function Field(ByVal rec As Variant, ByVal i As Long) As String
    Dim v As String
    If i < 0 Then Exit Function
    If i > UBound(rec) Then Exit Function
    v = Trim$(rec(i))
    If Len(v) > 1 And Left$(v, 1) = "'" And InStr("=+-@", Mid$(v, 2, 1)) > 0 Then v = Mid$(v, 2)
    Field = v
End Function
