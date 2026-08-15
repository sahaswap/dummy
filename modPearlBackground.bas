Option Explicit
'=====================================================================
' modPearlBackground - sets Sheet1's BACKGROUND image (behind cells).
' Two entry points, each embeds its own PNG (decoded with pure VBA -
' no chart, no clipboard, no COM):
'   ApplyPearlBackground   - dark near-solid navy background
'   ApplyLightBackground    - soft light pearl background
'   RemovePearlBackground   - clears it, gridlines back on
' Background pictures do not print, by design.
'=====================================================================
Private Const SHEET_NAME As String = "Sheet1"

Sub ApplyPearlBackground()
    ApplyBg DarkPngB64(), "Dark"
End Sub

Sub ApplyLightBackground()
    ApplyBg LightPngB64(), "Light"
End Sub

Private Sub ApplyBg(ByVal b64 As String, ByVal modeName As String)
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
    tmp = Environ$("TEMP") & "\bg_" & Format(Now, "hhmmss") & ".png"

    On Error GoTo Fail
    WriteBytesToFile tmp, Base64Decode(b64)
    If Dir(tmp) = "" Then Err.Raise 53, , "PNG was not written."

    ws.Activate
    ws.SetBackgroundPicture fileName:=tmp
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo Fail

    If wasProt Then ws.Protect Password:="p7ss"
    MsgBox modeName & " background applied. Run RemovePearlBackground to clear it.", _
           vbInformation, "Background"
    Exit Sub
Fail:
    Dim d As String: d = Err.Description
    On Error Resume Next
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
    MsgBox "Couldn't apply the background:" & vbCrLf & d, vbExclamation, "Background"
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
    Application.CommandBars.ExecuteMso "SheetBackgroundDelete"
    ActiveWindow.DisplayGridlines = True
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
    MsgBox "Background removed; gridlines restored.", vbInformation, "Background"
End Sub

' --- pure-VBA helpers ------------------------------------------------
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
                bitBuf = bitBuf Mod CLng(2 ^ bits)
                oi = oi + 1
            End If
        End If
    Next
    ReDim Preserve out(0 To oi - 1)
    Base64Decode = out
End Function

Private Function DarkPngB64() As String
    Dim b As String
    b = ""
    b = b & "iVBORw0KGgoAAAANSUhEUgAAAUAAAAFACAIAAABC8jL9AAAKjUlEQVR42u3dWa7sxhFFUX6qs2UJkCzbmv84Xe92rxr2TDLjAAvYU1jIn4yI4Yd//FG7P+f653g/tuzfh/p1ZX999VPb/nXrPw37+aH/ruq3g/2vVb9M9ft8f5dtoJdeekP13owM9NJLb6jeyoDppZfeBb1lAdNLL73LemsCppdeelfpLQiYXnrpXau3GmB66aV3g95SgOmll95teusAppdeejfrLQKYXnrp3aO3AmB66aV3p97ugOmll979evsCppdeeg/p7QiYXnrpPaq3F2B66aW3gd4ugOmll942eq8HTC+99DbTezFgeumlt6XeKwHTSy+9jfVeBpheeultr/cawPTSS+8pei8ATC+99J6l92zA9NJL74l6TwVML730nqv3PMD00kvv6XpPAkwvvfReofcMwPTSS+9FepsDppdeeq/T2xYwvfTSe6nehoDppZfeq/W2AkwvvfR20NsEML300ttH73HA9NJLbze9BwHTSy+9PfUeAUwvvfR21rsbML300ttf7z7A9NJLbwm9OwDTSy+9VfRuBUwvvfQW0rsJML300ltL73rA9NJLbzm9KwHTSy+9FfWuAUwvvfQW1bsImF566a2rdx4wvfTSW1rvDGB66aW3ut4pwPTSS2+A3lHA9NJLb4beV8D00ktvjN4nwPTSS2+S3nvA9NJLb5jeL8D00ktvnt53wPTSS2+k3iXA9NJLb+n+HOill95QvdOA6aWX3vJ6JwDTSy+9CXrHANNLL70hel8A00svvTl6HwHTSy+9UXrvANNLL71pej8B00svvYF63wDTSy+9mXonAdNLL7319d6oDvTSS2+o3hHA9NJLb4reZ8D00ktv"
    b = b & "kN4HwPTSS2+W3u+A6aWX3ji9H4DppZfeRL3fANNLL72hetsCppdeei/V+2M7wPTSS+/VelsBppdeejvobQKYXnrp7aP3OGB66aW3m96DgOmll96eeo8Appdeejvr3Q2YXnrp7a93H2B66aW3hN4dgOmll94qercCppdeegvp3QSYXnrpraV3PWB66aW3nN6VgOmll96KetcAppdeeovqXQRML7301tU7D5heeuktrXcGML300ltd7xRgeumlN0DvKGB66aU3Q+8rYHrppTdG7xNgeumlN0nvPWB66aU3TO8XYHrppTdP7ztgeumlN1LvYcD00ktvP703gwO99NIbqvcAYHrppbe33r2A6aWX3gJ6dwGml156a+jdDpheeukto3cjYHrppbeS3i2A6aWX3mJ6VwOml1566+ldB5heeuktqXcFYHrppbeq3iXA9NJLb2G9s4DppZfe2nqnAdNLL73l9U4AppdeehP0jgGml156Q/S+AKaXXnpz9D4CppdeeqP03gGml1560/R+AqaXXnoD9b4BppdeejP1rgdML730ltN7sznQSy+9oXrXAKaXXnqL6l0ETC+99NbVOw+YXnrpLa13BjC99NJbXe8UYHrppTdA7yhgeumlN0PvK2B66aU3Ru8TYHrppTdJ7z1geumlN0zvF2B66aU3T+87YHrppTdS7wNgeumlN0vvje1AL730hur9AEwvvfQm6v0GmF566Q3V2xowvfTSe6Hen1oCppdeeq/V2w4wvfTSe7neRoDppZfeHnpbAKaXXno76T0MmF566e2n9xhgeumlt6veA4DppZfe3nr3AqaXXnoL6N0FmF566a2hdztgeumlt4zejYDppZfeSnq3AKaXXnqL6V0NmF566a2ndx1geumlt6Te"
    b = b & "FYDppZfeqnqXANNLL72F9c4CppdeemvrnQZML730ltc7AZheeulN0DsGmF566Q3R+wKYXnrpzdH7CJheeumN0nsHmF566U3T+wmYXnrpDdT7BpheeunN1PsOmF566Y3Ue3t9B3rppTdU7w3dQC+99IbqbQaYXnrpvV5vG8D00ktvF70NANNLL7299B4FTC+99HbUewgwvfTS21fvfsD00ktvd707AdNLL70V9O4BTC+99BbRuxkwvfTSW0fvNsD00ktvKb0bANNLL73V9K4FTC+99BbUuwowvfTSW1PvMmB66aW3rN4FwPTSS29lvXOA6aWX3uJ6JwHTSy+99fWOA6aXXnoj9I4AppdeelP0PgOml156g/Q+AKaXXnqz9H4HTC+99Mbp/QBML730Jur9BpheeukN1fsEmF566U3S+/MdYHrppTdM7xdgeumlN0/vO2B66aU3Uu9qwPTSS289vTebA7300huqdwVgeumlt6reJcD00ktvYb2zgOmll97aeqcB00svveX1TgCml156E/SOAaaXXnpD9L4AppdeenP0PgKml156o/TeAaaXXnrT9H4CppdeegP1vgGml156M/UeB0wvvfR203szONBLL72heo8Appdeejvr3Q2YXnrp7a93H2B66aW3hN4dgOmll94qercCppdeegvp3QSYXnrpraV3PWB66aW3nN6VgOmll96KetcAppdeeovqXQRML7301tU7D5heeuktrXcGML300ltd7xRgeumlN0DvKGB66aU3Q+8rYHrppTdG7xNgeumlN0nvPWB66aU3TO8XYHrppTdP7ztgeumlN1JvS8D00kvvxXpv7gZ66aU3VG8bwPTSS28XvQ0A00svvb30HgVML730dtR7CDC99NLbV+9+wPTS"
    b = b & "S293vTsB00svvRX07gFML730FtG7GTC99NJbR+82wPTSS28pvRsA00svvdX0rgVML730FtS7CjC99NJbU+8yYHrppbes3gXA9NJLb2W9c4DppZfe4nonAdNLL7319Y4DppdeeiP0jgCml156U/Q+A6aXXnqD9D4AppdeerP0fgdML730xun9AEwvvfQm6v0GmF566Q3VOw2YXnrpLa/3l3HA9NJLb4LeMcD00ktviN4XwPTSS2+O3kfA9NJLb5TeO8D00ktvmt5PwPTSS2+g3jfA9NJLb6beRcD00ktvXb03pAO99NIbqncGML300ltd7xRgeumlN0DvKGB66aU3Q+8rYHrppTdG7xNgeumlN0nvPWB66aU3TO8XYHrppTdP7ztgeumlN1JvacD00kvv4k/ngV566Q3VWxQwvfTSu3LKaKCXXnpD9ZYDTC+99G6a8B3opZfeUL2FANNLL707tmsM9NJLb6jeEoDppZfe3ZutBnrppTdUb2fA9NJL78GtkgO99NIbqrcbYHrppbfJRueBXnrpDdXbATC99NLb8JrCQC+99IbqvRQwvfTS2/yS0UAvvfSG6r0IML300nvSFcGBXnrpDdV7OmB66aX31Au+A7300huq90TA9NJL79l6zwJML730XqD3FMD00kvvNXrbA6aXXnov09sYML300nul3paA6aWX3ov1NgNML730Xq+3DWB66aW3i94GgOmll95eeo8CppdeejvqPQSYXnrp7at3P2B66aW3u96dgOmll94KevcAppdeeovo3QyYXnrpraN3G2B66aW3lN4NgOmll95qetcCppdeegvqXQWYXnrpral3GTC99NJbVu8CYHrppbey3jnA9NJLb3G9k4DppZfe+nrHAdNLL70RekcA00sv"
    b = b & "vSl6nwHTSy+9QXofANNLL71Zer8DppdeeuP0fgCml156E/Xe+j99OBqTSy9FUwAAAABJRU5ErkJggg=="
    DarkPngB64 = b
End Function

Private Function LightPngB64() As String
    Dim b As String
    b = ""
    b = b & "iVBORw0KGgoAAAANSUhEUgAAAWgAAAFoCAIAAAD1h/aCAAANJ0lEQVR42u3cSW5VyQJF0TP/uSWFa2NckfUfxX8PGwwkUuaG7pKOdKXbiEY0Vi/2frn637/u1Ze7/Gqvv7c3T7v4+/Pe/mMn59/u9OudnX278y93+tdhF9/b5Tc7ednVT+70eZenf361s5ddnP1xcf71Lo47v/j9uMuXnV399rLr406vfz3u3afdHHdy8+G49y97e/v4srvj3tw9HHf/9R4eXj/cv+zxq716vHvehx/Z6+Nuv9mbx6e9P+zty25Ovtjp47vjHo47e7h+2vlxV4ddPO3+8vLTru4vnnZ9f/60d4fdnR1287zT9592e3dy3O1xd7dvP+/+uDdPe3je68Me3/9zrz7kfXPC8QIenvf2afc3J592eve8s9vnnR/2/ubi4y5vnnf17rB314ddP33ePf16f9jlcbcXx90d9vFGDlfzcHbc42Gnx3047OT614/77bC3x/1+2Jvj/nja66s/P+2vw1497+/P++W475gwalCDGtRIavw7HNSgBjWo0eCgBjWoQY0GBzWoQQ1qNDioQQ1qUKPBQQ1qUIMaDQ5qUIMa1GhwUIMa1KBGg4Ma1KAGNRoc1KAGNajR4KAGNahBjQYHNahBDWr8BBzUoAY1qPHfXr2OGtSgBjWSGp/goAY1qEGNUtgYNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5"
    b = b & "Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjaufgoMa1KAGNRoc1KAGNajR4KAGNahBjQYHNahBDWo0OKhBDWpQ48fgoAY1qEGN//5WftSgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1"
    b = b & "yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1Kiv5EcNalCDGhkOalCDGtT4ATioQQ1qUKN1eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1"
    b = b & "qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyj"
    b = b & "BjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXLM2pQgxrUqF2eUYMa1KBG7fKMGtSgBjVql2fUoAY1qFG7PKMGNahBjdrlGTWoQQ1q1C7PqEENalCjdnlGDWpQgxq1yzNqUIMa1KhdnlGDGtSgRu3yjBrUoAY1apdn1KAGNahRuzyjBjWoQY3a5Rk1qEENatQuz6hBDWpQo3Z5Rg1qUIMatcszalCDGtSoXZ5RgxrUoEbt8owa1KAGNWqXZ9SgBjWoUbs8owY1qEGN"
    b = b & "2uUZNahBDWrULs+oQQ1qUKN2eUYNalCDGrXL83/zCxcSjEtqRQAAAABJRU5ErkJggg=="
    LightPngB64 = b
End Function
