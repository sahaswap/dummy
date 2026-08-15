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
    b = b & "iVBORw0KGgoAAAANSUhEUgAAAeAAAAHgCAIAAADytinCAAAX+UlEQVR42u3Wx3Zc57WF0f0Eki1bIpWoYNkKzDkBJMGcJMuWLaf7/s9x93/OKVSBAkmwg1qNOcbXXO05Vr333ZXD+/6Q3v9hv6vvn97ozLrfnV11btX57lr3+wurLq66tPTB5bnrH1xZdXXpD9c2ur7Rje7G3B9vHtatd+n26MPRzXU7Sx91u92tpTu3Tux393Z3srs3t3Nyb+fjufujT+7vfvJg1cPdTx/eGT0afdY97u6Ontz9fO7pvblT3bP99k493/tivxeb3f9y7uWR+mruxejrFw/m/vR89M3cs4d/nvpL9/Tht08ffdc9efT91A+PH5/uHj0+8+jx2UdPzj58cu7hk/MPuqcXHjy9eP/ppW7v2eW9Z1e6e8+v3nt+7e7z66MXN+68uNntvri1+/J2t/NyZ+fl7u0fuzvdrR/v3vrx3q2f7t38aW/qfnfjr92DqYfXlx5d/3l0bfT4yE37vy1dXXq49Pe5B3NXul/ur9ob/WPv8uje0j/vzl3qfr2zanf0r92Lo53Rv7vbcxe6/9xadXP035vnRzfW/a+7vt+5/f7v2rt0dfS/uSsbXT7/37lLo/9cnLow+nd3furchX91Z0e/nrnw6+nRP3+Y+v7iP7rvRr98e/GXv4z+/uepby79rfvT6OevL/381eivX059cfmn7tTSj59PfTZ6+emqT6682O/jdc9PvkPPRpefnVh6utGTE5dGH40eL118NPfhxYejCw/m/njh/uj83qp7fzjX3V1154Oz3e7SmZ3u92duL52+1f3u9M2lH26suv7+3PfdtY2udu91380d4nDRmc50pjOdA3V+DdB0pjOd6Uznbet8GNB0pjOd"
    b = b & "6UznAJ1/AzSd6UxnOtM5Q+eDQNOZznSmM51jdN4Ams50pjOd6Zyk8wpoOtOZznSmc5jOE9B0pjOd6UznPJ0PB5rOdKYznem8dZ1b46IznelMZzoH6vwq0HSmM53pTOcQnQ8ATWc605nOdM7ReQ00nelMZzrTOUrnBWg605nOdKZzms4DaDrTmc50pnOgzptA05nOdKYznYN0bpmLznSmM53pHKjzDDSd6UxnOtM5TucJaDrTmc50pnOezv2ei850pjOd6Ryoc7NcdKYznelM50CdV0DTmc50pjOdw3SegKYznelMZzrn6XwAaDrTmc50pnOOzs1y0ZnOdKYznQN1XoCmM53pTGc6p+k8gKYznelMZzoH6rwBNJ3pTGc60zlJ55a56ExnOtOZzoE6T0DTmc50pjOd83TeAJrOdKYznemcpHPLXHSmM53pTOdAnSeg6UxnOtOZznk6z0DTmc50pjOd43Tu91x0pjOd6UznQJ0b56IznelMZzoH6rwBNJ3pTGc60zlJ55a56ExnOtOZzoE6T0DTmc50pjOd83TeAJrOdKYznemcpHPLXHSmM53pTOdAnSeg6UxnOtOZznk6r4GmM53pTGc6R+ncMhed6UxnOtM5UOcBNJ3pTGc60zlQ532g6UxnOtOZzlk6N85FZzrTmc50DtR5AprOdKYznemcp3O/56IznelMZzoH6twyF53pTGc60zlQ5wloOtOZznSmc57Oa6DpTGc605nOUTq3zEVnOtOZznQO1HkATWc605nOdA7U+SDQdKYznelM5xidm+WiM53pTGc6B+q8AprOdKYznekcpvMENJ3pTGc60zlP54NA05nOdKYznWN0bpaLznSmM53pHKjzCmg605nOdKZzmM4T0HSmM53pTOc8nWeg"
    b = b & "6UxnOtOZznE693suOtOZznSmc6DOLXPRmc50pjOdA3UeQNOZznSmM50DdX4N0HSmM53pTOdt69waF53pTGc60zlQ598ATWc605nOdM7Q+SDQdKYznelM5xidN4CmM53pTGc6J+m8AprOdKYznekcpvMENJ3pTGc60zlP53cEms50pjOd6XxcOre6RWc605nOdA7U+chA05nOdKYznY9X56MBTWc605nOdD52nY8ANJ3pTGc603kbOr8NaDrTmc50pvOWdH4j0HSmM53pTOft6fx6oOlMZzrTmc5b1fk1QNOZznSmM523rfNhQNOZznSmM50DdP4N0HSmM53pTOcMnQ8CTWc605nOdI7ReQNoOtOZznSmc5LOK6DpTGc605nOYTpPQNOZznSmM53zdF6ApjOd6UxnOqfp3DgXnelMZzrTOVDnGWg605nOdKZznM4fHgCaznSmM53pHKNzs1x0pjOd6UznQJ1XQNOZznSmM53DdJ6ApjOd6UxnOufpvAaaznSmM53pHKVzy1x0pjOd6UznQJ0H0HSmM53pTOdAnSeg6UxnOtOZznk6fzQBTWc605nOdI7Tud9z0ZnOdKYznQN1bpmLznSmM53pHKjzBDSd6UxnOtM5T+cBNJ3pTGc60zlQ537PRWc605nOdA7U+cQCNJ3pTGc60zlM5wloOtOZznSmc57OC9B0pjOd6UznNJ0b56IznelMZzoH6jwBTWc605nOdM7T+eQKaDrTmc50pnOWzo1z0ZnOdKYznQN1noCmM53pTGc65+nc77noTGc605nOgTr3ey4605nOdKZzoM4fD6DpTGc605nOeTovQNOZznSmM53TdG6ci850pjOd6Ryo8wQ0nelMZzrTOU/nfs9FZzrTmc50DtS5ZS4605nO"
    b = b & "dKZzoM4T0HSmM53pTOc8nQfQdKYznelM50Cd+z0XnelMZzrTOVDnxrnoTGc605nOgTovQNOZznSmM53TdG6ci850pjOd6Ryo8wQ0nelMZzrTOU/nzyag6UxnOtOZznE693suOtOZznSmc6DOjXPRmc50pjOdA3UeQNOZznSmM50Dde73XHSmM53pTOdAnT8fQNOZznSmM53zdF4DTWc605nOdI7SuWUuOtOZznSmc6DOA2g605nOdKZzoM4T0HSmM53pTOc8nU9tAE1nOtOZznQO0rllLjrTmc50pnOgzjPQdKYznelM5zidJ6DpTGc605nOeTr3ey4605nOdKZzoM79novOdKYznekcqPMXC9B0pjOd6UznMJ0noOlMZzrTmc55Or8CNJ3pTGc60zlF52a56ExnOtOZzoE67wNNZzrTmc50ztJ5BprOdKYznekcp/MAms50pjOd6Ryoc7/nojOd6UxnOgfq/OUAms50pjOd6Zyn81GBpjOd6UxnOh+zzm1v0ZnOdKYznQN1fjvQdKYznelM563o/Bag6UxnOtOZztvS+U1A05nOdKYznbeo82uBpjOd6UxnOm9X58OBpjOd6UxnOm9d50OApjOd6UxnOifo/CrQdKYznelM5xCdDwBNZzrTmc50ztF5DTSd6UxnOtM5SucFaDrTmc50pnOazgNoOtOZznSmc6DOK6DpTGc605nOYTp/NYCmM53pTGc65+m8AE1nOtOZznRO07lxLjrTmc50pnOgzmug6UxnOtOZzlE6t8xFZzrTmc50DtR5AE1nOtOZznQO1HkBms50pjOd6Zymc+NcdKYznelM50CdV0DTmc50pjOdw3T+ZgBNZzrTmc50ztN5AE1nOtOZznQO1Lnfc9GZznSmM50DdW6c"
    b = b & "i850pjOd6Ryo8wQ0nelMZzrTOU/nfs9FZzrTmc50DtS533PRmc50pjOdA3Xu91x0pjOd6UznQJ37PRed6UxnOtM5UOd+z0VnOtOZznQO1LlxLjrTmc50pnOgzgNoOtOZznSmc6DO/Z6LznSmM53pHKhzv+eiM53pTGc6B+rc77noTGc605nOgTqfGUDTmc50pjOd83Tu91x0pjOd6UznQJ37PRed6UxnOtM5UOdzE9B0pjOd6UznOJ37PRed6UxnOtM5UOd+z0VnOtOZznQO1Lnfc9GZznSmM50Dde73XHSmM53pTOdAnfs9F53pTGc60zlQ537PRWc605nOdA7U+fIAms50pjOd6Zync7/nojOd6UxnOgfq3O+56ExnOtOZzoE693suOtOZznSmc6DO/Z6LznSmM53pHKhzv+eiM53pTGc6B+rc77noTGc605nOgTr3ey4605nOdKZzoM79novOdKYznekcqHO/56IznelMZzoH6tzvuehMZzrTmc6BOvd7LjrTmc50pnOgzv2ei850pjOd6Ryo884MNJ3pTGc60zlN58a56ExnOtOZzoE6T0DTmc50pjOd83Tu91x0pjOd6UznQJ37PRed6UxnOtM5UOe7A2g605nOdKZzns79novOdKYznekcqHO/56IznelMZzoH6tw4F53pTGc60zlQ5wloOtOZznSmc57O92eg6UxnOtOZzmk6N85FZzrTmc50DtR5AZrOdKYznemcpnPjXHSmM53pTOdAnddA05nOdKYznaN0bpmLznSmM53pHKjzAJrOdKYznekcqPMKaDrTmc50pnOYzo1z0ZnOdKYznQN1XoCmM53pTGc6p+nc3had6UxnOtM5UOd3AJrOdKYznel8nDo/PiLQdKYznelM52PW"
    b = b & "+UhA05nOdKYznY9f57cDTWc605nOdN6Kzm8Bms50pjOd6bwtnd8ENJ3pTGc603mLOr8WaDrTmc50pvN2dT4caDrTmc50pvPWdT4EaDrTmc50pnOCzq8CTWc605nOdA7R+QDQdKYznelM5xyd10DTmc50pjOdo3RegKYznelMZzqn6TyApjOd6UxnOgfqPANNZzrTmc50jtP50RpoOtOZznSmc5LOLXPRmc50pjOdA3WegKYznelMZzrn6bwGms50pjOd6Rylc8tcdKYznelM50CdB9B0pjOd6UznQJ33gaYznelMZzpn6fxwAprOdKYznekcp/MaaDrTmc50pnOUzi1z0ZnOdKYznQN1HkDTmc50pjOdA3VeAU1nOtOZznQO0/nBAJrOdKYznemcp/MMNJ3pTGc60zlO537PRWc605nOdA7Uud9z0ZnOdKYznQN1vj8DTWc605nOdE7TeQBNZzrTmc50DtR5BprOdKYznekcp/PeAJrOdKYznemcp3O/56IznelMZzoH6tw4F53pTGc60zlQ532g6UxnOtOZzlk635uApjOd6UxnOsfpPICmM53pTGc6B+rc77noTGc605nOgTrfHUDTmc50pjOd83SegaYznelMZzrH6dzvuehMZzrTmc6BOvd7LjrTmc50pnOgzndmoOlMZzrTmc5pOg+g6UxnOtOZzoE6z0DTmc50pjOd43TeHUDTmc50pjOd83Tu91x0pjOd6UznQJ0b56IznelMZzoH6jwDTWc605nOdI7TeWcGms50pjOd6Zymc+NcdKYznelM50CdV0DTmc50pjOdw3S+PYCmM53pTGc65+k8A01nOtOZznSO07nfc9GZznSmM50Dde73XHSmM53pTOdAnW/NQNOZznSmM53TdB5A"
    b = b & "05nOdKYznQN1noGmM53pTGc6x+l8cwBNZzrTmc50ztO533PRmc50pjOdA3VunIvOdKYznekcqPMm0HSmM53pTOcgnW+sgKYznelMZzpn6TwDTWc605nOdI7TeQGaznSmM53pnKZz41x0pjOd6UznQJ03gKYznelMZzon6Xx9AZrOdKYznekcpvMENJ3pTGc60zlP502g6UxnOtOZzkE6t8xFZzrTmc50DtR5BprOdKYznekcp/MAms50pjOd6Ryoc7/nojOd6UxnOgfqfO3oQNOZznSmM52PU+ejAk1nOtOZznQ+Zp2PBDSd6UxnOtP5+HV+O9B0pjOd6Uznrej8FqDpTGc605nO29L5TUDTmc50pjOdt6jza4GmM53pTGc6b1fnw4GmM53pTGc6b13nQ4CmM53pTGc6J+j8KtB0pjOd6UznEJ0PAE1nOtOZznTO0XkNNJ3pTGc60zlK5wVoOtOZznSmc5rOA2g605nOdKZzoM4z0HSmM53pTOc4na/uA01nOtOZznSO0rllLjrTmc50pnOgzgNoOtOZznSmc6DOB4CmM53pTGc65+h8ZR9oOtOZznSmc5TOC9B0pjOd6UznNJ0H0HSmM53pTOdAnddA05nOdKYznaN0bpmLznSmM53pHKjzAJrOdKYznekcqPMMNJ3pTGc60zlO537PRWc605nOdA7Uud9z0ZnOdKYznQN1bpyLznSmM53pHKjzDDSd6UxnOtM5TucLM9B0pjOd6UznNJ0b56IznelMZzoH6rwATWc605nOdE7TuXEuOtOZznSmc6DOC9B0pjOd6UznNJ0b56IznelMZzoH6jwDTWc605nOdI7Tud9z0ZnOdKYznQN17vdcdKYznelM50Cdz0xA05nOdKYzneN07vdcdKYz"
    b = b & "nelM50Cd+z0XnelMZzrTOVDnxrnoTGc605nOgTovQNOZznSmM53TdG6ci850pjOd6Ryo8ww0nelMZzrTOU7nfs9FZzrTmc50DtS533PRmc50pjOdA3X+dgKaznSmM53pHKdzv+eiM53pTGc6B+rc77noTGc605nOgTo3zkVnOtOZznQO1HkBms50pjOd6Zymc+NcdKYznelM50CdZ6DpTGc605nOcTr3ey4605nOdKZzoM79novOdKYznekcqPPXE9B0pjOd6UznOJ37PRed6UxnOtM5UOd+z0VnOtOZznQO1LlxLjrTmc50pnOgzgvQdKYznelM5zSdG+eiM53pTGc6B+q8DzSd6UxnOtM5S+dTE9B0pjOd6UznOJ0H0HSmM53pTOdAnfs9F53pTGc60zlQ58a56ExnOtOZzoE6z0DTmc50pjOd43Tu91x0pjOd6UznQJ37PRed6UxnOtM5UOdPZ6DpTGc605nOaToPoOlMZzrTmc6BOh8Ams50pjOd6Zyjc7NcdKYznelM50CdF6DpTGc605nOaToPoOlMZzrTmc6BOm8CTWc605nOdA7S+eMV0HSmM53pTOcsnWeg6UxnOtOZznE6D6DpTGc605nOgTr3ey4605nOdKZzoM4njww0nelMZzrT+Vh1PiLQdKYznelM5+PW+ShA05nOdKYznbeg81uBpjOd6UxnOm9H5zcDTWc605nOdN6azm8Ams50pjOd6bxNnV8HNJ3pTGc603nLOh8KNJ3pTGc603n7Ov8WaDrTmc50pnOEzq8ATWc605nOdE7ReRNoOtOZznSmc5DO+0DTmc50pjOds3SegaYznelMZzrH6bwCms50pjOd6Rymc+NcdKYznelM50CdB9B0pjOd6UznQJ37PRed6Uxn"
    b = b & "OtM5UOcTE9B0pjOd6UznOJ1fAZrOdKYznemconOzXHSmM53pTOdAnfeBpjOd6UxnOmfpPANNZzrTmc50jtN5AprOdKYznemcp3O/56IznelMZzoH6tw4F53pTGc60zlQ5xloOtOZznSmc5zOH62BpjOd6UxnOifp3DIXnelMZzrTOVDnCWg605nOdKZzns5roOlMZzrTmc5ROrfMRWc605nOdA7UeQBNZzrTmc50DtR5BTSd6UxnOtM5TOfGuehMZzrTmc6BOq+BpjOd6UxnOkfp3DIXnelMZzrTOVDnATSd6UxnOtM5UOcV0HSmM53pTOcwnRvnojOd6UxnOgfqvAk0nelMZzrTOUjnlrnoTGc605nOgTrPQNOZznSmM53jdB5A05nOdKYznQN17vdcdKYznelM50Cd+z0XnelMZzrTOVDnlrnoTGc605nOgTrPQNOZznSmM53jdB5A05nOdKYznQN17vdcdKYznelM50Cd+z0XnelMZzrTOVDnlrnoTGc605nOgTpPQNOZznSmM53zdF6ApjOd6UxnOqfp3DgXnelMZzrTOVDnDaDpTGc605nOSTq3zEVnOtOZznQO1HkCms50pjOd6Zyn8wI0nelMZzrTOU3nxrnoTGc605nOgTpvAE1nOtOZznRO0rllLjrTmc50pnOgzhPQdKYznelM5zydN4GmM53pTGc6B+ncMhed6UxnOtM5UOcZaDrTmc50pnOczgNoOtOZznSmc6DO/Z6LznSmM53pHKjz+wNoOtOZznSmc57OM9B0pjOd6UznOJ37PRed6UxnOtM5UOdmuehMZzrTmc6BOu8DTWc605nOdM7SeQaaznSmM53pHKfzAjSd6UxnOtM5TefGuehMZzrTmc6BOk9A05nOdKYznfN0"
    b = b & "fm8FNJ3pTGc60zlL58a56ExnOtOZzoE6vyPQdKYznelM5+PSudUtOtOZznSmc6DORwaaznSmM53pfLw6Hw1oOtOZznSm87HrfASg6UxnOtOZztvQ+W1A05nOdKYznbekc/f/jA+g72Us3NgAAAAASUVORK5CYII="
    PearlPngB64 = b
End Function
