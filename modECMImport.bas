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
' Two ways in:
'   AUTOMATIC  The workbook watches Downloads (started on open by Auto_Open).
'              When a new ECM_Alert_Info csv appears it moves it to
'              Downloads\ECM_Imported and runs steps 1-4 for every alert in
'              it, one after another, with no prompts. Sheet1 is left on the
'              last alert. A new export is only picked up while the form is
'              empty or untouched since the last import (no Decision, Template
'              or counterparties entered) - otherwise the status bar asks for
'              Reset, so work in progress is never overwritten.
'              StopEcmWatcher / StartEcmWatcher turn it off and on.
'   MANUAL     START with an empty ECM ID, or Alt+F8 > Import_ECM_Info:
'              pick the csv and the alert.
' ==========================================================================

Private Const PWD As String = "p7ss"
Private Const MAX_LISTED As Long = 15
Private Const AUTO_EXPORT_MODE As String = "EN"   ' "EN" / "LEGACY" / "PIVOT" / "" = do not export
Private Const WATCH_SECONDS As Long = 10          ' how often Downloads is checked
Private Const NEXT_ALERT_SECONDS As Long = 5      ' gap between alerts in a batch
Private Const IMPORTED_FOLDER As String = "ECM_Imported"

' Loaded csv
Private mRows As Collection
Private mCEcm As Long, mCVel As Long, mCCust As Long, mCAddr As Long, mCCountry As Long, mCFile As Long

' Watcher / batch state
Private mWatching As Boolean
Private mNextTick As Date
Private mBatchCsv As String
Private mBatchNext As Long
Private mBatchLog As String
Private mBatchOk As Long
Private mBatchFailed As Long

' ==========================================================================
' MANUAL
' ==========================================================================
Public Sub Import_ECM_Info()
    Dim ws As Worksheet, csvPath As String, loadErr As String
    Dim pick As Long, ecmID As String, note As String, copied As String

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
        RunAutoExport False
        Application.StatusBar = False
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
' AUTOMATIC - Downloads watcher
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
    Application.StatusBar = "ECM auto-import is on - watching Downloads for ECM exports."
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
    On Error Resume Next      ' a failed check must never stop the watcher
    CheckDownloads
    On Error GoTo 0
    If mWatching Then ScheduleTick
End Sub

Private Sub CheckDownloads()
    Dim csv As String
    If mBatchCsv <> "" Then Exit Sub          ' a batch is already running
    csv = WaitingCsv(DownloadsFolder())
    If csv = "" Then Exit Sub
    If FormInUse(ThisWorkbook.Sheets("Sheet1")) Then
        Application.StatusBar = "New ECM export in Downloads - press Reset to import it " & _
                                "(case " & ThisWorkbook.Sheets("Sheet1").Range("J9").Value & " is in progress)."
    Else
        StartBatch csv
    End If
End Sub

' Work has started on the case on Sheet1 (Decision, Template or a counterparty
' entered). An empty form, or one only filled by an import, is not "in use".
Private Function FormInUse(ByVal ws As Worksheet) As Boolean
    Dim r As Long
    If Trim$(CStr(ws.Range("J9").Value)) = "" Then Exit Function
    If Trim$(CStr(ws.Range("J5").Value)) <> "" Or Trim$(CStr(ws.Range("J6").Value)) <> "" Then FormInUse = True: Exit Function
    For r = 18 To 23
        If Trim$(CStr(ws.Cells(r, "J").Value)) <> "" Then FormInUse = True: Exit Function
    Next r
End Function

' Oldest ECM_Alert_Info csv sitting directly in Downloads that has finished
' saving (unchanged for 3 seconds). Processed files are moved out, so anything
' here is new.
Private Function WaitingCsv(ByVal folder As String) As String
    Dim f As String, best As String, bestT As Date, t As Date
    On Error Resume Next
    f = Dir$(folder & "\ECM_Alert_Info_*.csv")
    Do While f <> ""
        t = FileDateTime(folder & "\" & f)
        If Now - t > TimeSerial(0, 0, 3) Then
            If best = "" Or t < bestT Then best = folder & "\" & f: bestT = t
        End If
        f = Dir$()
    Loop
    WaitingCsv = best
End Function

Private Sub StartBatch(ByVal csv As String)
    Dim fso As Object, destFolder As String, dest As String, loadErr As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Move it out of Downloads first so it is only ever picked up once.
    destFolder = DownloadsFolder() & "\" & IMPORTED_FOLDER
    If Not fso.FolderExists(destFolder) Then fso.CreateFolder destFolder
    dest = destFolder & "\" & fso.GetFileName(csv)
    If fso.FileExists(dest) Then dest = destFolder & "\" & fso.GetBaseName(csv) & "_" & Format$(Now, "hhnnss") & ".csv"
    fso.MoveFile csv, dest

    loadErr = LoadCsv(dest)
    If loadErr <> "" Then
        Application.StatusBar = "ECM auto-import skipped " & fso.GetFileName(dest) & ": " & loadErr
        Exit Sub
    End If

    mBatchCsv = dest
    mBatchNext = 2
    mBatchLog = ""
    mBatchOk = 0
    mBatchFailed = 0
    Application.StatusBar = "ECM auto-import: " & (mRows.Count - 1) & " alert(s) found - starting..."
    Application.OnTime Now + TimeSerial(0, 0, 1), MacroRef("ECM_Batch_Next")
End Sub

' One alert per call, then schedules itself for the next one. The gap lets the
' export's deferred tracker update run while Sheet1 still shows that alert.
Public Sub ECM_Batch_Next()
    Dim ws As Worksheet, idx As Long, total As Long, ecmID As String
    Dim note As String, copied As String, result As String

    If mBatchCsv = "" Then Exit Sub
    On Error GoTo Failed
    If LoadCsv(mBatchCsv) <> "" Then GoTo Finish

    idx = mBatchNext
    total = mRows.Count - 1
    If idx > mRows.Count Then GoTo Finish
    ecmID = Field(mRows(idx), mCEcm)

    ThisWorkbook.Activate
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Activate
    Application.StatusBar = "ECM auto-import: alert " & (idx - 1) & " of " & total & " (ECM " & ecmID & ")..."

    copied = ImportRow(ws, idx, mBatchCsv, note)
    If copied = "" Then
        result = "transaction file not found in Downloads"
    ElseIf AUTO_EXPORT_MODE = "" Then
        result = "OK"
    Else
        result = RunAutoExport(True)
        If result = "" Then result = "export did not finish (see its message)"
    End If
    GoTo Record

Failed:
    result = "error " & Err.Number & ": " & Err.Description
    ProtectSheet ThisWorkbook.Sheets("Sheet1")
    Application.EnableEvents = True
    Resume Record

Record:
    If result = "OK" Then mBatchOk = mBatchOk + 1 Else mBatchFailed = mBatchFailed + 1
    mBatchLog = mBatchLog & ecmID & vbTab & result & vbCrLf
    mBatchNext = idx + 1
    If mBatchNext <= mRows.Count Then
        Application.OnTime Now + TimeSerial(0, 0, NEXT_ALERT_SECONDS), MacroRef("ECM_Batch_Next")
        Exit Sub
    End If

Finish:
    FinishBatch
End Sub

Private Sub FinishBatch()
    Dim logPath As String, msg As String, ts As Object
    On Error Resume Next
    logPath = Left$(mBatchCsv, Len(mBatchCsv) - 4) & "_import_log.txt"
    Set ts = CreateObject("Scripting.FileSystemObject").CreateTextFile(logPath, True)
    ts.Write "ECM ID" & vbTab & "Result" & vbCrLf & mBatchLog
    ts.Close
    On Error GoTo 0

    msg = "ECM auto-import finished: " & mBatchOk & " OK" & IIf(mBatchFailed > 0, ", " & mBatchFailed & " failed", "") & _
          ". Sheet1 shows ECM " & ThisWorkbook.Sheets("Sheet1").Range("J9").Value & "."
    Application.StatusBar = msg
    If mBatchFailed > 0 Then
        MsgBox msg & vbCrLf & vbCrLf & mBatchLog & vbCrLf & "Log: " & logPath, vbExclamation, "ECM auto-import"
    End If
    mBatchCsv = ""
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
' (batch mode). Returns "OK" when the export finished, "" otherwise.
Private Function RunAutoExport(ByVal quiet As Boolean) As String
    gExportLastResult = ""
    gExportQuiet = quiet
    gExportModeOverride = AUTO_EXPORT_MODE
    Consolidated_AML_Workflow
    gExportModeOverride = ""
    gExportQuiet = False
    RunAutoExport = gExportLastResult
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
    Dim folder As String, newest As String, ans As VbMsgBoxResult
    Dim picked As Variant

    folder = DownloadsFolder()
    newest = NewestCsv(folder)
    If newest = "" Then newest = NewestCsv(folder & "\" & IMPORTED_FOLDER)
    If newest <> "" And NewestCsv(folder & "\" & IMPORTED_FOLDER) <> "" Then
        If FileDateTime(NewestCsv(folder & "\" & IMPORTED_FOLDER)) > FileDateTime(newest) Then _
            newest = NewestCsv(folder & "\" & IMPORTED_FOLDER)
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
    f = Dir$(folder & "\ECM_Alert_Info_*.csv")
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
Private Function MakeCaseFolders(ByVal ecmID As String) As String
    Dim fso As Object, userProfile As String, desktopPath As String, mainFolder As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    userProfile = Environ$("USERPROFILE")
    If fso.FolderExists(userProfile & "\OneDrive - Community Federal Savings Bank\Desktop") Then
        desktopPath = userProfile & "\OneDrive - Community Federal Savings Bank\Desktop"
    Else
        desktopPath = userProfile & "\Desktop"
    End If
    mainFolder = desktopPath & "\" & ecmID
    If Not fso.FolderExists(mainFolder) Then fso.CreateFolder mainFolder
    If Not fso.FolderExists(mainFolder & "\Transaction Files") Then fso.CreateFolder mainFolder & "\Transaction Files"
    If Not fso.FolderExists(mainFolder & "\Pivot") Then fso.CreateFolder mainFolder & "\Pivot"
    MakeCaseFolders = mainFolder
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

    fso.CopyFile src, destFolder & "\", True
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
