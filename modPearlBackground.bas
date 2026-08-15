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
    b = b & "iVBORw0KGgoAAAANSUhEUgAAAggAAAIICAIAAACEnoE0AAAXg0lEQVR42u3ZWW4cxxqE0Xi3LM+WRM2TZ/vuf3e3CRrOqO7qiaQfXHmAWMN/kPnli2fvL93zk3ux2JMXH8Zu1vZyba+O7PXYl68/jr05t7ePvXeLPV3d++P78PGrY/u4sq93+3R8n1f2zd5+eoR9e7tPt/s89t3Y592+/9T76Ye7fbzdj7f7+W7PPtztl92e3+39Ly/e//rPbnZ799vdXt7u991evR17/faP27253Zvb/fnP3r6+2193e7e3V7v97/2D9+7k3r5c35uDvV7bq5v9vVzbzZG9eHFqzx9jz8b+2tuPK/vzh91uFvv+7/1xt++W+/Z2v/e+Gfttt6/39vK3r273696eruyXu335mPt5sZvV/bTbk/V9fvJidZ9W9vxuH/f2xdiH/T3b26XXPlSgAhWoQAUqXA8DFahABSpQYQ4VLoOBClSgAhWoMI0KF8BABSpQgQpUmEmFczBQgQpUoAIVJlPhJAxUoAIVqECF+VQ4DgMVqEAFKlBhShWOwEAFKlCBClSYVYU1GKhABSpQgQoTq3AAAxWoQAUqUGFuFZYwUIEKVKACFaZXoWCgAhWoQAUqUGHAQAUqUIEKVKDCgIEKVKACFahAhUthoAIVqEAFKkymwu74hwpUoAIVqECF7s2hAhWoQAUqUKF7c6hABSpQgQpUaAVCBSpQgQpUoMJxGKhABSpQgQpzq7CEgQpUoAIVqDC9CgUDFahABSpQgQoDBipQgQpUoAIVBgxUoAIVqEAFKhyDgQpUoAIVqDC5CjsLQgUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdFYIFahABSpQgQpHYKACFahABSpMr8KOg1CB"
    b = b & "ClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdG8OFahABSpQgQr9SAgVqEAFKlCBCgcwUIEKVKACFagwYKACFahABSpQ4QwMVKACFahAhVlV2CkQKlCBClSgAhW6N4cKVKACFahAhe7NoQIVqEAFKlChHwmhAhWoQAUqUOEABipQgQpUoAIVBgxUoAIVqEAFKpyHgQpUoAIVqDClCjsCQgUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqtAWhAhWoQAUqUGENBipQgQpUoAIVBgxUoAIVqEAFKgwYqEAFKlCBClRYhYEKVKACFahAhR0HoQIVqEAFKlChe3OoQAUqUIEKVOjeHCpQgQpUoAIVujeHClSgAhWoQIXuzaECFahABSpQoXtzqEAFKlCBClTo3hwqUIEKVKACFbo3hwpUoAIVqECF7s2hAhWoQAUqUKF7c6hABSpQgQpU6N4cKlCBClSgAhWagFCBClSgAhWocBIGKlCBClSgwsQqHMBABSpQgQpUmFuFJQxUoAIVqECF6VUoGKhABSpQgQpUGDBQgQpUoAIVqDBgoAIVqEAFKlDhX4SBClSgAhWo8F9WYXfJQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUq9EkPFahABSpQgQqPCgMVqEAFKlBhQyo8GAYqUIEKVKDCtlR4GAxUoAIVqECFzanwABioQAUqUIEKW1ThvjBQgQpUoAIVNqrCvWCgAhWoQAUqbFeF62GgAhWoQAUqbFqFK2GgAhWoQAUqbF2Fa2CgAhWoQAUqTKDCxTBQgQpUoAIV5lDhMhioQAUqUIEK06hwAQxUoAIVqECFmVQ4BwMVqEAF"
    b = b & "KlBhMhVOwkAFKlCBClSYT4XjMFCBClSgAhWmVOEIDFSgAhWoQIVZVViDgQpUoAIVqDCxCgcwUIEKVKACFeZWYQkDFahABSpQYXoVCgYqUIEKVKACFQYMVKACFahABSoMGKhABSpQgQpUOAYDFahABSpQYXIVdhaEClSgAhWoQIXuzaECFahABSpQoXtzqEAFKlCBClTorBAqUIEKVKACFc7BQAUqUIEKVJhVhacrMFCBClSgAhUmVuEABipQgQpUoMLcKixhoAIVqEAFKkyvQsFABSpQgQpUoMKAgQpUoAIVqECFAQMVqEAFKlCBChfBQAUqUIEKVJhPhd39DxWoQAUqUIEK3ZtDBSpQgQpUoEL35lCBClSgAhWo0BCEClSgAhWoQIUjMFCBClSgAhWmV6FgoAIVqEAFKlBhwEAFKlCBClSgwoCBClSgAhWoQIWGgQpUoAIVqECF7s2hAhWoQAUqUKF7c6hABSpQgQpU6N4cKlCBClSgAhW6N4cKVKACFahAha/WYaACFahABSpMr0LBQAUqUIEKVKDCgIEKVKACFahAhQEDFahABSpQgQqnYaACFahABSpMq8JOgVCBClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCv1ICBWoQAUqUIEK+zBQgQpUoAIVqDBgoAIVqEAFKlBhCQMVqEAFKlCBCtWbQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdG8OFahABSpQgQoNQahABSpQgQpUOAIDFahABSpQYXoVCgYqUIEKVKACFQYMVKACFahABSoMGKhABSpQgQpUOA0DFahABSpQYVoVdgqEClSgAhWoQIXuzaECFahABSpQoXtzqEAFKlCBClToR0KoQAUqUIEKVNiHgQpUoAIV"
    b = b & "qECFAQMVqEAFKlCBCsdhoAIVqEAFKsytwjcLGKhABSpQgQrTq1AwUIEKVKACFagwYKACFahABSpQYcBABSpQgQpUoMLjwkAFKlCBClTYjAq7qx4qUIEKVKACFbo3hwpUoAIVqECF7s2hAhWoQAUqUKHPe6hABSpQgQpUeCgMVKACFahAha2qcB8YqEAFKlCBChtW4WoYqEAFKlCBCttW4ToYqEAFKlCBCptX4QoYqEAFKlCBCjOocCkMVKACFahAhUlUuAgGKlCBClSgwjwqnIeBClSgAhWoMJUKZ2CgAhWoQAUqzKbCKRioQAUqUIEKE6pwFAYqUIEKVKDCnCqsw0AFKlCBClSYVoUVGKhABSpQgQozq7APAxWoQAUqUGFyFRYwUIEKVKACFagwYKACFahABSpQYcBABSpQgQpUoMKAgQpUoAIVqECFPRioQAUqUIEKVBi9OVSgAhWoQAUqdG8OFahABSpQgQrdm0MFKlCBClSgQvfmUIEKVKACFajQvTlUoAIVqEAFKnRvDhWoQAUqUIEK3ZtDBSpQgQpUoEL35lCBClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdG8OFahABSpQgQrdm0MFKlCBClSgQvfmUIEKVKACFajQvTlUoAIVqEAFKnRvDhWoQAUqUIEK3ZtDBSpQgQpUoEL35lCBClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdG8OFahABSpQgQrdm0MFKlCBClSgQvfmUIEKVKACFajQvTlUoAIVqEAFKnRvDhWoQAUqUIEK3ZtDBSpQgQpUoEL3"
    b = b & "5lCBClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdG8OFahABSpQgQrdm0MFKlCBClSgQvfmUIEKVKACFajQvTlUoAIVqEAFKnRvDhWoQAUqUIEK3ZtDBSpQgQpUoEL35lCBClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdG8OFahABSpQgQrdm0MFKlCBClSgQvfmUIEKVKACFajQvTlUoAIVqEAFKnRvDhWoQAUqUIEK3ZtDBSpQgQpUoEL35lCBClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdG8OFahABSpQgQrdm0MFKlCBClSgQvfmUIEKVKACFajwbgEDFahABSpQgQpLGKhABSpQgQpUGL05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqvH8gDFSgAhWoQIWtqnAfGKhABSpQgQobVuFqGKhABSpQgQrbVuE6GKhABSpQgQqbV+EKGKhABSpQgQozqHApDFSgAhWoQIVJVLgIBipQgQpUoMI8KpyHgQpUoAIVqDCVCmdgoAIVqEAFKsymwikYqEAFKlCBChOqcBQGKlCBClSgwpwqrMNABSpQgQpUmFaFFRioQAUqUIEKM6uwDwMVqEAFKlBhchUWMFCBClSgAhWoMGCgAhWoQAUqUGHAQAUqUIEKVKDCgIEKVKACFahAhUthoAIVqEAFKsymwrsTMFCBClSgAhUmVOEoDFSgAhWoQIU5VViH"
    b = b & "gQpUoAIVqDCtCiswUIEKVKACFWZWYR8GKlCBClSgwuQqLGCgAhWoQAUqUGHAQAUqUIEKVKDCgIEKVKACFahAhQEDFahABSpQgQrnYaACFahABSrMqcKOgFCBClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCm1BqEAFKlCBClRYgYEKVKACFahAhQEDFahABSpQgQoDBipQgQpUoAIVTsFABSpQgQpUmFmFN3swUIEKVKACFSZXYQEDFahABSpQgQoDBipQgQpUoAIVBgxUoAIVqEAFKgwYqEAFKlCBClQ4AwMVqEAFKlBhWhVeH8JABSpQgQpUmFmFfRioQAUqUIEKk6uwgIEKVKACFahAhQEDFahABSpQgQoDBipQgQpUoAIVBgxUoAIVqEAFKpyCgQpUoAIVqDCzCjsIQgUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUq9O9RqEAFKlCBClRYwEAFKlCBClSgwhkYqEAFKlCBCtOq8PIQBipQgQpUoMLMKuzDQAUqUIEKVJhchQUMVKACFahABSoMGKhABSpQgQpUGDBQgQpUoAIVqDBgoAIVqEAFKlDhPAxUoAIVqECFOVW4WYWBClSgAhWoMK0KKzBQgQpUoAIVZlZhHwYqUIEKVKDC5CosYKACFahABSpQYcBABSpQgQpUoMKAgQpUoAIVqECFAQMVqEAFKlCBCpfCQAUqUIEKVJhNhd3xDxWoQAUqUIEK3ZtDBSpQgQpUoEL35lCBClSgAhWo0AqEClSgAhWoQIWjMFCBClSgAhUmV2EBAxWoQAUqUIEKAwYqUIEKVKACFQYMVKACFahABSoMGKhABSpQgQpUeGQYqEAFKlCBCptR4fnDYaACFahABSpsSYWHwkAFKlCBClTY"
    b = b & "mAoPgoEKVKACFaiwPRXuDwMVqEAFKlBhkyrcEwYqUIEKVKDCVlW4DwxUoAIVqECFDatwNQxUoAIVqECFbatwHQxUoAIVqECFzatwBQxUoAIVqECFGVS4FAYqUIEKVKDCJCpcBAMVqEAFKlBhHhXOw0AFKlCBClSYSoUzMFCBClSgAhVmU+EUDFSgAhWoQIUJVTgKAxWoQAUqUGFOFdZhoAIVqEAFKkyrwgoMVKACFahAhZlV2IeBClSgAhWoMLkKCxioQAUqUIEKVBgwUIEKVKACFagwYKACFahABSpQYcBABSpQgQpUoMIqDFSgAhWoQAUq3IoQKlCBClSgAhW6N4cKVKACFahAhe7NoQIVqEAFKlChe3OoQAUqUIEKVOjeHCpQgQpUoAIVujeHClSgAhWoQIX+PQoVqEAFKlCBCgsYqEAFKlCBClQ4DQMVqEAFKlBhXhV+PICBClSgAhWoMLUKezBQgQpUoAIVZlehYaACFahABSpQYcBABSpQgQpUoMKAgQpUoAIVqECFgoEKVKACFahAhe7NoQIVqEAFKlChe3OoQAUqUIEKVOjeHCpQgQpUoAIVujeHClSgAhWoQIXuzaECFahABSpQobNCqEAFKlCBClQ4hIEKVKACFahAhb/LQqhABSpQgQpU6N4cKlCBClSgAhW6N4cKVKACFahAhe7NoQIVqEAFKlChe3OoQAUqUIEKVOjeHCpQgQpUoAIVujeHClSgAhWoQIXuzaECFahABSpQobNCqEAFKlCBClTYg4EKVKACFahAhdGbQwUqUIEKVKBC9+ZQgQpUoAIVqNC9OVSgAhWoQAUqdG8OFahABSpQgQqdFUIFKlCBClSgwioMVKACFahABSrcihAqUIEKVKACFbo3hwpUoAIVqECF"
    b = b & "7s2hAhWoQAUqUKF7c6hABSpQgQpU6N4cKlCBClSgAhW6N4cKVKACFahAhe7NoQIVqEAFKlChe3OoQAUqUIEKVPh6AQMVqEAFKlCBCg0DFahABSpQgQrdm0MFKlCBClSgQvfmUIEKVKACFajQvTlUoAIVqEAFKnRvDhWoQAUqUIEK3ZtDBSpQgQpUoEL/HoUKVKACFahAhQUMVKACFahABSqchoEKVKACFagwrwpPD2CgAhWoQAUqTK3CHgxUoAIVqECF2VVoGKhABSpQgQpUGDBQgQpUoAIVqDBgoAIVqEAFKlDhAAYqUIEKVKACFe7+kEIFKlCBClSgQvfmUIEKVKACFajQvTlUoAIVqEAFKnRvDhWoQAUqUIEK3ZtDBSpQgQpUoEL35lCBClSgAhWo8OXjwUAFKlCBClTYlAoPhIEKVKACFaiwNRUeAgMVqEAFKlBhgyrcGwYqUIEKVKDCNlW4HwxUoAIVqECFzapwDxioQAUqUIEKW1bhWhioQAUqUIEKG1fhKhioQAUqUIEK21fhchioQAUqUIEKU6hwIQxUoAIVqECFWVS4BAYqUIEKVKDCRCqchYEKVKACFagwlwqnYaACFahABSpMp8IJGKhABSpQgQozqnAMBipQgQpUoMKkKqzCQAUqUIEKVJhXhUMYqEAFKlCBClOrsAcDFahABSpQYXYVGgYqUIEKVKACFQYMVKACFahABSoMGKhABSpQgQpUOAYDFahABSpQYW4VdhaEClSgAhWoQIXuzaECFahABSpQoXtzqEAFKlCBClTorBAqUIEKVKACFc7CQAUqUIEKVJhUhZ0CoQIVqEAFKlChe3OoQAUqUIEKVOjeHCpQgQpUoAIV+pEQKlCBClSgAhUOYaACFahABSpQYcBABSpQ"
    b = b & "gQpUoMISBipQgQpUoAIV/ikLoQIVqEAFKlChe3OoQAUqUIEKVOjeHCpQgQpUoAIVujeHClSgAhWoQIXuzaECFahABSpQ4ckaDFSgAhWoQAUqDBioQAUqUIEKVBgwUIEKVKACFajQMFCBClSgAhWoUL05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC9+ZQgQpUoAIVqNCPhFCBClSgAhWocAgDFahABSpQgQoDBipQgQpUoAIVTsNABSpQgQpUmFWFnQKhAhWoQAUqUKF7c6hABSpQgQpU6N4cKlCBClSgAhX6kRAqUIEKVKACFQ5goAIVqEAFKlBhwEAFKlCBClSgwgEMVKACFahABSr8XRZCBSpQgQpUoEL35lCBClSgAhWo0L05VKACFahABSp0bw4VqEAFKlCBCt2bQwUqUIEKVKBC/x6FClSgAhWoQIUFDFSgAhWoQAUqrMJABSpQgQpUoMKtCKECFahABSpQoXtzqEAFKlCBClTo3hwqUIEKVKACFbo3hwpUoAIVqECF7s2hAhWoQAUqUKF7c6hABSpQgQpU6N+jUIEKVKACFaiwhIEKVKACFahAheMwUIEKVKACFaZWYWdBqEAFKlCBClTo3hwqUIEKVKACFbo3hwpUoAIVqECFzgqhAhWoQAUqUOFfhoEKVKACFajwn1Vhd8ZDBSpQgQpUoEL35lCBClSgAhWo0L05VKACFahABSr0PQ8VqEAFKlCBCo8NAxWoQAUqUGErKjwGDFSgAhWoQIUNqfBgGKhABSpQgQrbUuFhMFCBClSgAhU2p8IDYKACFahABSpsUYX7wkAFKlCBClTYqAr3goEKVKACFaiwXRWuh4EKVKACFaiwaRV2+z+q/61Mh1yRWQAAAABJRU5ErkJg"
    b = b & "gg=="
    PearlPngB64 = b
End Function
