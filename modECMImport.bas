Attribute VB_Name = "modECMImport"
Option Explicit

' ==========================================================================
' IMPORT ECM INFO
'
' Fills Sheet1 from the ECM_Alert_Info_<date>_<time>.csv that the "ECM Export"
' browser button saves to Downloads (one row per alert):
'   J9  = ECM Case           <- ECM ID
'   J10 = Alert ID           <- Velocity Alert ID
'   J13 = Customer Name      <- Customer Name
'   J14 = Customer Address   <- Address (street, city, state zip, country)
'   L27 = ISO Code 2         <- country code of the address; Sheet1's own
'                               change handler then fills G27 (Country Name)
'                               and O27 (ISO Code 3), and the risk formulas follow.
' Cells are written one at a time with events on, so Sheet1 applies its
' usual auto-capitalisation exactly as if they had been typed.
'
' Then creates the case folders (same rule as START) and copies the alert's
' transaction file from Downloads into <Desktop>\<ECM>\Transaction Files,
' ready for "Export Trx File".
'
' Run it from START (when the ECM ID is empty) or with Alt+F8 > Import_ECM_Info.
' ==========================================================================

Private Const PWD As String = "p7ss"
Private Const MAX_LISTED As Long = 15

Public Sub Import_ECM_Info()
    Dim ws As Worksheet
    Dim csvPath As String, txt As String
    Dim rows As Collection, hdr As Variant, rec As Variant
    Dim cEcm As Long, cVel As Long, cCust As Long, cAddr As Long, cCountry As Long, cFile As Long
    Dim pick As Long, ecmID As String, velID As String, custName As String
    Dim addr As String, iso2 As String, trxFile As String
    Dim caseFolder As String, copied As String, summary As String
    Dim prevEvents As Boolean

    Set ws = ThisWorkbook.Sheets("Sheet1")
    prevEvents = Application.EnableEvents

    ' ---- 1. Which CSV ----
    csvPath = ChooseCsv()
    If csvPath = "" Then Exit Sub

    On Error GoTo Fail
    txt = ReadUtf8(csvPath)
    Set rows = ParseCsv(txt)
    If rows.Count < 2 Then
        MsgBox "That file has no alert rows:" & vbCrLf & csvPath, vbExclamation, "Import ECM Info"
        Exit Sub
    End If

    hdr = rows(1)
    cEcm = ColIdx(hdr, "ECM ID")
    cVel = ColIdx(hdr, "Velocity Alert ID")
    cCust = ColIdx(hdr, "Customer Name")
    cAddr = ColIdx(hdr, "Address")
    cCountry = ColIdx(hdr, "Country")
    cFile = ColIdx(hdr, "Transaction File")
    If cEcm < 0 Then
        MsgBox "This does not look like an ECM export file (no 'ECM ID' column):" & vbCrLf & csvPath, _
               vbExclamation, "Import ECM Info"
        Exit Sub
    End If

    ' ---- 2. Which alert ----
    pick = ChooseAlert(rows, cEcm, cVel, cCust, Trim$(CStr(ws.Range("J9").Value)))
    If pick < 2 Then Exit Sub
    rec = rows(pick)

    ecmID = Field(rec, cEcm)
    velID = Field(rec, cVel)
    custName = Field(rec, cCust)
    addr = Field(rec, cAddr)
    iso2 = CountryCode(Field(rec, cCountry))
    trxFile = Field(rec, cFile)

    If Trim$(CStr(ws.Range("J9").Value)) <> "" And Trim$(CStr(ws.Range("J9").Value)) <> ecmID Then
        If MsgBox("Sheet1 already holds ECM case " & ws.Range("J9").Value & "." & vbCrLf & vbCrLf & _
                  "Replace the case details with ECM case " & ecmID & "?" & vbCrLf & _
                  "(Tip: press Reset first to clear the rest of the form.)", _
                  vbYesNo + vbQuestion, "Import ECM Info") <> vbYes Then Exit Sub
    End If

    ' ---- 3. Fill Sheet1 ----
    Application.EnableEvents = True
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
    Application.EnableEvents = prevEvents

    summary = "Filled Sheet1 for ECM case " & ecmID & ":" & vbCrLf & _
              "   ECM Case (J9), Alert ID (J10)" & _
              IIf(custName <> "", ", Customer Name (J13)", "") & _
              IIf(addr <> "", ", Customer Address (J14)", "") & vbCrLf
    If iso2 <> "" Then
        If Trim$(CStr(ws.Range("G27").Value)) <> "" Then
            summary = summary & "   Country (G27): " & ws.Range("G27").Value & vbCrLf
        Else
            summary = summary & "   Country code " & iso2 & " is not in the country table - pick the country in G27." & vbCrLf
        End If
    End If

    ' ---- 4. Folders + transaction file ----
    caseFolder = MakeCaseFolders(ecmID)
    copied = CopyTransactionFile(csvPath, trxFile, ecmID, caseFolder & "\Transaction Files")
    If copied <> "" Then
        summary = summary & vbCrLf & "Transaction file copied to:" & vbCrLf & "   " & copied & vbCrLf
    Else
        summary = summary & vbCrLf & "Transaction file NOT found in Downloads - copy it into:" & vbCrLf & _
                  "   " & caseFolder & "\Transaction Files" & vbCrLf
    End If

    summary = summary & vbCrLf & "Next: choose Decision and Template, then click Export Trx File."
    MsgBox summary, vbInformation, "Import ECM Info"
    Exit Sub

Fail:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = prevEvents
    On Error GoTo 0
    MsgBox "Import stopped with an error:" & vbCrLf & vbCrLf & "Err " & errNum & ": " & errDesc, _
           vbCritical, "Import ECM Info"
End Sub

' --------------------------------------------------------------------------
' Newest ECM_Alert_Info_*.csv in Downloads (with a Yes/No to use it), or a
' file picker. Returns "" if cancelled.
' --------------------------------------------------------------------------
Private Function ChooseCsv() As String
    Dim folder As String, newest As String, ans As VbMsgBoxResult
    Dim picked As Variant

    folder = Environ$("USERPROFILE") & "\Downloads"
    newest = NewestCsv(folder)

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
' "<ECM ID>_*.xls*" next to the CSV) into destFolder. Returns the new path, or "".
Private Function CopyTransactionFile(ByVal csvPath As String, ByVal trxFile As String, _
                                     ByVal ecmID As String, ByVal destFolder As String) As String
    Dim fso As Object, srcFolder As String, src As String, f As String
    Dim bestT As Date, t As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    srcFolder = Left$(csvPath, InStrRev(csvPath, "\") - 1)

    If trxFile <> "" Then
        If fso.FileExists(srcFolder & "\" & trxFile) Then src = srcFolder & "\" & trxFile
    End If
    If src = "" Then
        f = Dir$(srcFolder & "\" & ecmID & "_*.xls*")
        Do While f <> ""
            t = FileDateTime(srcFolder & "\" & f)
            If src = "" Or t > bestT Then src = srcFolder & "\" & f: bestT = t
            f = Dir$()
        Loop
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
