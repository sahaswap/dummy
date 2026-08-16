Option Explicit
'=====================================================================
' modIconDiag - Segoe MDL2 icon tester.
'
' IconDiagnostic  - builds a throwaway "_IconTest" sheet: each row is
'                   an icon we need, with 4 candidate glyph codes shown
'                   in Segoe MDL2 Assets. Photograph it and tell me which
'                   code (the small hex next to each glyph) is the RIGHT
'                   picture for each row - then I lock those into the theme.
' RemoveIconTest  - deletes the _IconTest sheet when you're done.
'=====================================================================

Sub IconDiagnostic()
    Dim wbProt As Boolean
    On Error Resume Next
    wbProt = ThisWorkbook.ProtectStructure
    ThisWorkbook.Unprotect Password:="p7ss"
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("_IconTest").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "_IconTest"
    ws.Activate

    ws.Range("A1").Value = "Segoe MDL2 icon test - tell me which CODE is the right icon for each row"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "(each glyph has its hex code shown just to its left)"

    Dim data As Variant
    data = Array( _
        Array("Play (Start)",         "E768", "E102", "E71B", "EC57"), _
        Array("Upload (Export Trx)",  "E898", "E74A", "EA35", "E78C"), _
        Array("Search (OSDD)",        "E721", "E094", "E11A", "EB51"), _
        Array("Edit (Gen Narrative)", "E70F", "E104", "E70E", "E90F"), _
        Array("Rename",               "E8AC", "E13E", "EB7E", "E71C"), _
        Array("Document (PDF Merge)", "E8A5", "EA90", "E7C3", "E160"), _
        Array("Warm-Up (power/fire)", "E945", "E7E8", "E706", "EC49"), _
        Array("Refresh (Reset)",      "E72C", "E117", "E777", "E895"), _
        Array("Alert (warn/shield)",  "E7BA", "E730", "E814", "EA18"), _
        Array("Contact (Customer)",   "E77B", "E13D", "E748", "EA8C"), _
        Array("People (Counterparty)", "E716", "E902", "EF58", "E7EE"), _
        Array("Globe (Country)",      "E774", "E128", "E12B", "E909"))

    Dim r As Long, i As Long, j As Long, col As Long, code As String
    r = 4
    For i = LBound(data) To UBound(data)
        ws.Cells(r, 1).Value = data(i)(0)
        ws.Cells(r, 1).Font.Bold = True
        col = 2
        For j = 1 To 4
            code = CStr(data(i)(j))
            ws.Cells(r, col).Value = code
            ws.Cells(r, col).Font.Size = 9
            ws.Cells(r, col + 1).Value = GlyphOf(code)
            ws.Cells(r, col + 1).Font.Name = "Segoe MDL2 Assets"
            ws.Cells(r, col + 1).Font.Size = 22
            ws.Cells(r, col + 1).HorizontalAlignment = xlCenter
            col = col + 2
        Next j
        r = r + 2
    Next i

    ws.Columns.AutoFit
    ws.Range("A1").Select

    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    MsgBox "Created the '_IconTest' sheet." & vbCrLf & vbCrLf & _
           "Photograph it and tell me, per row, which hex code is the correct icon " & _
           "(e.g. 'Play = E768, Rename = the 2nd one'). Then run RemoveIconTest to delete it.", _
           vbInformation, "Icon Diagnostic"
End Sub

Sub RemoveIconTest()
    Dim wbProt As Boolean
    On Error Resume Next
    wbProt = ThisWorkbook.ProtectStructure
    ThisWorkbook.Unprotect Password:="p7ss"
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("_IconTest").Delete
    Application.DisplayAlerts = True
    If wbProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    On Error GoTo 0
End Sub

Private Function GlyphOf(ByVal code As String) As String
    On Error Resume Next
    GlyphOf = ChrW(CLng(Val("&H" & code)))
    On Error GoTo 0
End Function
