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
' The image is a seamless 200x200 tile, so Excel repeats it across the
' sheet as soft diagonal iridescent bands (deep nacre: magenta -> gold
' -> green -> teal -> violet -> rose -> blue -> back).
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
    b = b & "iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAIAAAAiOjnJAAAHaUlEQVR42u3cV1pVBwBF4T3E2EsssUSNmiMWlKKAdPjwCiJCECVosGCnXemXLgEJWJAmiPhCdwg5gzj7bc3h/9bj0krK65W0FytXG1Yznqxm1a9mP1jNrV3Lr1krrF4rrlwrub0eK1svja2XlayXF21U5G9U5mz8lbVxN2PzXvpmbcpmXfLmwwtb9Ulbj//cenp669nJny+O/3x1dPP1kZU3x340/r7YdGK2+dTnltMTLWdHW4OheFLi7YWOtkut7ZfftF953pH6uDP9Qde1e92ZVT3Xy3tyYom8ot6C3L7CzP7itIGS5MFY0mDpmaFbJ4ZvH3l357fhqsPD1YeGag4N1R4crDs48PDAQP2B/ie/9jXs73u+v/flvt43+xJNe3ta9vbE93S37+nq3N3Vvaszsaujb2fH4M724R3tIzvaRre/Hd/+dmJb/MMvrZ+2tUztaPqy+/WXvS+n9z+bPvh05vCjmaP/zB6vmz1ZO/dHzdzZ6vmgcv58xdeLt74m3/yacmMhrXjhWuFiZt5ids633KxvBRlLRelLJanfY5e/l15aLj+/XHHuR1WwXB0s3QsW/w4WHgTz9cHs42C6IfjyIvj8KvjYGEw2BxPxYLw9GOsM/u0J3vUGwwPB4FDQPxL0jgWJ8aB7Muj8GLRPBW0zQXwuaFkImheDxrmk5qkLrR8uxt8nt41c6RhI7UqkdXdeTcQz+pqyBl5mDzXkDNfnjdQVjNYUjlUVj5eXTNyMfSgu/ZRf9vl6+ZdrFTMpd2YvVc2fq144U/PtxP3vR2uXj9Uun7q/dLZm8fzdhcvV8+lVs1mV07kVU0W3P8XKP5aXTVaW/lcTe193Y/RR"
    b = b & "yciz4nevioaaCwba8vu68xL9OT3D2V1jWR2TmW1TGfG5qy2L6c0/UhuFKlRFriqslVCFqshVhbUSqlAVuaqwVkIVqiJXFdZKqEJV5KrCWglVqIpcVVgroQpVkasKayVUoSpyVWGthCpURa4qrJVQharIVYW1EqpQFbmqsFZCFaoiVxXWSqhCVeSqwloJVaiKXFVYK6EKVZGrCmslVKEqclVhrYQqVEWuKqyVUIWqyFWFtRKqUBW5qrBWQhWqIlcV1kqoQlXkqsJaCVWoilxVWCuhClWRqwprJVShKnJVYa2EKlRFriqslVCFqshVhbUSqlAVuaqwVkIVqhz/DqEKVY5/h1CFKse/Q6hClePfIVShyvHvEKpQ5fh3CFWocvw7hCpUOf4dQhWqHP8OoQpVjn+HUIUqx79DqEKV498hVKHK8e8QqlDl+HcIVahy/DuEKlQ5/h1CFaoc/w6hClWOf4dQhSrHv0OoQpXj3yFUocrx7xCqUOX4dwhVqHL8O4QqVDn+HUIVqhz/DqEKVY5/h1CFKse/Q6hClePfIVShyvHvEKpQ5fh3CFWocvw7hCpUOf4dQhWqHP8OoQpVjn+HUIUqx79DqEKV498hVKHK8e8QqlDl+HcIVahy/DuEKlQ5/h1CFaoc/w6hClWOf4dQhSrHv0OoQpXj3yFUocrx7xCqUOX4dwhVqHL8O4QqVDn+HUIVqhz/DqEKVY5/h1CFKse/Q6hClePfIVShyvHvEKpQ5fh3CFWocvw7hCpUOf4dQhWqHP8OoQpVjn+HUIUqx79DqEKV498hVKHK8e8QqlDl+HcIVahy/DuEKlQ5/h1CFaoc/w6hClWOf4dQhSrHv0OoQpXj3yFUocrx7xCqUOX4dwhVqHL8O4QqVDn+HUIVqhz/"
    b = b & "DqEKVY5/h1CFKse/Q6hClePfIVShyvHvEKpQ5fh3CFWocvw7hCpUOf4dQhWqHP8OoQpVjn+HUIUqx79DqEKV498hVKHK8e8QqlDl+HcIVahy/DuEKlQ5/h1CFaoc/w6hClWOf4dQhSrHv0OoQpXj3yFUocrx7xCqUOX4dwhVqHL8O4QqVDn+HUIVqhz/DqEKVY5/h1CFKse/Q6hClePfIVShyvHvEKpQ5fh3CFWocvw7hCpUOf4dQhWqHP8OoQpVjn+HUIUqx79DqEKV498hVKHK8e8QqlDl+HcIVahy/DuEKlQ5/h1CFaoc/w6hClWOf4dQhSrHv0OoQpXj3yFUocrx7xCqUOX4dwhVqHL8O4QqVDn+HUIVqhz/DqEKVY5/h1CFKse/Q6hClePfIVShyvHvEKpQ5fh3CFWocvw7hCpUOf4dQhWqHP8OoQpVjn+HUIUqx79DqEKV498hVKHK8e8QqlDl+HcIVahy/DuEKlQ5/h1CFaoc/w6hClWOf4dQhSrHv0OoQpXj3yFUocrx7xCqUOX4dwhVqHL8O4QqVDn+HUIVqhz/DqEKVY5/h1CFKse/Q6hClePfIVShyvHvEKpQ5fh3CFWocvw7hCpUOf4dQhWqHP8OoQpVjn+HUIUqx79DqEKV498hVKHK8e8QqlDl+HcIVahy/DuEKlQ5/h1CFaoc/w6hClWOf4dQhSrHv0OoQpXj3yFUocrx7xCqUOX4dwhVqHL8O4QqVDn+HUIVqhz/DqEKVY5/h1CFKse/Q6hClePf8T94NZvdAVyAHQAAAABJRU5ErkJggg=="
    PearlPngB64 = b
End Function
