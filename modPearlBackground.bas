Option Explicit
'=====================================================================
' modPearlBackground - mother-of-pearl gradient as Sheet1's BACKGROUND.
'
' The pearl gradient PNG is EMBEDDED in this module as base64 text and
' written straight to disk with plain VBA (no chart, no clipboard, no
' COM objects). The old chart-export route produced a blank white image
' on this machine (the clipboard->chart paste silently dropped the
' picture); this route physically cannot come out white.
'
' The image is a seamless 360x360 tile - a MINIMALIST frosted-glass wash:
' near-white with only the faintest cool shimmer (whisper blue / lilac /
' mint / warm-white), broad and soft so it reads as frosted glass, not
' a rainbow.
'
' It touches NO cells, rows, values, formulas, dropdowns or buttons.
'
' Caveats (both fine for a working dashboard):
'   - Background pictures do NOT print.
'   - It's static: a flat sheet has no viewing angle, so it can't
'     actually shift color like real nacre - it just looks pearly.
'
'   ApplyPearlBackground  - decode + write PNG + set it as the background
'   RemovePearlBackground - clear it and turn gridlines back on
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"

Sub ApplyPearlBackground()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo 0

    Dim tmp As String
    tmp = Environ$("TEMP") & "\pearl_bg_" & Format(Now, "hhmmss") & ".png"

    On Error GoTo Fail
    WriteBytesToFile tmp, Base64Decode(PearlPngB64())
    If Dir(tmp) = "" Then Err.Raise 53, , "PNG was not written."

    ws.Activate
    ws.SetBackgroundPicture fileName:=tmp
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo Fail

    If wasProt Then ws.Protect Password:="p7ss"
    MsgBox "Mother-of-pearl background applied to " & SHEET_NAME & _
           " (behind your data)." & vbCrLf & vbCrLf & _
           "Run RemovePearlBackground to clear it. It won't print, by design.", _
           vbInformation, "Pearl Background"
    Exit Sub

Fail:
    Dim d As String: d = Err.Description
    On Error Resume Next
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
    MsgBox "Couldn't apply the pearl background:" & vbCrLf & d, vbExclamation, "Pearl Background"
End Sub

Sub RemovePearlBackground()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    ws.Activate
    Application.CommandBars.ExecuteMso "SheetBackgroundDelete"   ' clears bg
    ActiveWindow.DisplayGridlines = True
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
    MsgBox "Pearl background removed; gridlines restored.", vbInformation, "Pearl Background"
End Sub

' --- pure-VBA helpers (no COM, cannot be blocked) -------------------

Private Sub WriteBytesToFile(ByVal path As String, ByRef bytes() As Byte)
    Dim f As Integer
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, bytes
    Close #f
End Sub

Private Function Base64Decode(ByVal s As String) As Byte()
    Static tbl(255) As Long
    Static ready As Boolean
    Dim i As Long, alpha As String
    If Not ready Then
        For i = 0 To 255: tbl(i) = -1: Next
        alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        For i = 1 To Len(alpha): tbl(Asc(Mid$(alpha, i, 1))) = i - 1: Next
        ready = True
    End If

    Dim out() As Byte
    ReDim out(0 To (Len(s) * 3) \ 4)
    Dim bitBuf As Long, bits As Long, oi As Long, ch As Long
    For i = 1 To Len(s)
        ch = tbl(Asc(Mid$(s, i, 1)))
        If ch >= 0 Then
            bitBuf = bitBuf * 64 + ch
            bits = bits + 6
            If bits >= 8 Then
                bits = bits - 8
                out(oi) = (bitBuf \ CLng(2 ^ bits)) And 255
                bitBuf = bitBuf Mod CLng(2 ^ bits)   ' drop consumed high bits (else bitBuf overflows Long)
                oi = oi + 1
            End If
        End If
    Next
    ReDim Preserve out(0 To oi - 1)
    Base64Decode = out
End Function

Private Function PearlPngB64() As String
    Dim b As String
    b = ""
    b = b & "iVBORw0KGgoAAAANSUhEUgAAAWgAAAFoCAIAAAD1h/aCAAAMJ0lEQVR42u3cwU4bCRRE0fr/vw1gA1lNRhOJhAD27ciz8ZFqaW9a9lm09O6+nb5/toeje+x7evqx1wM79Z0P73TVTqeXt52P7rnv5bc9vTxfu9cjO127t2dzPraX83Pfy/udrtzr1fvlW+cP9nx+vrzLnzi92x8/ladLe/y51y/28N9O7/b2p36Hw6hBDWpQI6nxKRzUoAY1qPGZGh/DQQ1qUIMaX6jxARzUoAY1qPG1Gu/hoAY1qEGNi2r8Bgc1qEENalyjxhsc1KAGNahxpRo/4aAGNahBjevV+BcOalCDGtRIahyHgxrUoMbdqvFwDA5qUIMa96zGETioQQ1q3LkaGQ5qUIMa1GhwUIMa1KBGg4Ma1KAGNRoc1KAGNajR4KAGNahBjQYHNahBDWo0OKhBDWpQo8FBDWpQgxoNDmpQgxrUaHBQgxrUoEaDgxrUoAY1GhzUoAY1qNHgoAY1qEGNBgc1qEENajQ4qEENalCjwUENalCDGjeHgxrUoMadq/FY4aAGNahBjQYHNahBDWo0OKhBDWpQo8FBDWpQgxoNDmpQgxrUaHBQgxrUoEaDgxrUoAY1GhzUoAY1qNHgoAY1qEGNBgc1qEENajQ4qEENalCjwUENalCDGg0OalCDGtRocFCDGtSgRoODGtSgBjUaHNSgBjWocQAOalCDGtRoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1"
    b = b & "yzNqUIMa1KhX8qMGNahBjQwHNahBDWrcHA5qUIMad67GqcJBDWpQgxoNDmpQgxrUaHBQgxrUoEaDgxrUoAY1GhzUoAY1qNHgoAY1qEGNBgc1qEENajQ4qEENalCjwUENalCDGg0OalCDGtRocFCDGtSgRoODGtSgBjUaHNSgBjWo0eCgBjWoQY0GBzWoQQ1qNDioQQ1qUOP/goMa1KDGvapxPggHNahBjTtW4xAc1KAGNe5bjQ4HNahBjbtXI8JBDWpQgxoNDmpQgxrUaHBQgxrUoEaDgxrUoAY1GhzUoAY1qNHgoAY1qEGNBgc1qEENajQ4qEENalCjwUENalCDGg0OalCDGtRocFCDGtSgRoODGtSgBjUaHNSgBjWo0eCgBjWoQY0GBzWoQQ1q3AIOalCDGtT49X3oqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatRj11GDGtSgxiE4qEENalCjwUENalCDGjeHgxrUoMZ9q/HDgVGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCD"
    b = b & "GtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGjWvMWpQgxrU+Cs4qEENalDjmvu1UYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY16pX8qEENalAjw0ENalCDGjeCgxrUoAY13t6HjhrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoMb5CBzUoAY1qNHgoAY1qEGNCgc1"
    b = b & "qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1LzGqEENalDjL+GgBjWoQY3L92ujBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG"
    b = b & "7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnn+AY1DOvtGbfxdAAAAAElFTkSuQmCC"
    PearlPngB64 = b
End Function
