Attribute VB_Name = "Module4"
Option Explicit

'====================================================================
'  SEARCH MATRIX - Raw URL / Open URL columns
'
'  Column C = "Raw URL"  -> plain text, built by worksheet formula
'                           (unchanged, no length limit issue here)
'  Column E = "Open URL" (header) -> each cell shows "Link", built as
'                           a REAL Excel Hyperlink object by this
'                           macro, styled with Excel's built-in
'                           "Hyperlink" cell style (blue, underlined).
'
'  WHY THIS HAS TO BE VBA:
'  Excel's HYPERLINK() worksheet FUNCTION silently caps its link
'  argument at 255 characters and returns #VALUE! past that - which
'  is exactly what was happening on the Negative News rows, since
'  those search-query URLs run ~330 characters once all the
'  arrest/corruption/laundering/etc. keywords are concatenated.
'  A genuine Hyperlink object (Range.Hyperlinks.Add, used below) has
'  no such 255-char restriction, so this sidesteps the error entirely.
'
'  AUTO-REFRESH:
'  This Sub itself just rebuilds the links once, on demand. The part
'  that makes it happen automatically whenever Sheet1 changes lives
'  in Sheet1's own code module - see the second code block below,
'  which you paste in separately (sheet-module code can't be
'  delivered via a plain .bas import).
'====================================================================
Sub RefreshSearchMatrixHyperlinks()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim rawURL As String

    On Error GoTo CleanFail

    Set ws = ThisWorkbook.Sheets("Search Matrix")

    ' last data row, based on the "Type of Search" column
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' header stays "Open URL" - only the per-row display text is "Link"
    ws.Range("E1").Value = "Open URL"

    ' wipe any existing hyperlink objects + text before rebuilding
    ws.Range("E2:E" & lastRow).ClearContents
    On Error Resume Next
    ws.Range("E2:E" & lastRow).Hyperlinks.Delete
    On Error GoTo CleanFail

    For r = 2 To lastRow
        rawURL = CStr(ws.Range("C" & r).Value)   ' Raw URL column - already plain text

        ' 1) Manage the link + its font (underline/colour only).
        If Len(rawURL) > 0 Then
            ' Real Hyperlink object - no 255-char cap like HYPERLINK() has
            ws.Hyperlinks.Add Anchor:=ws.Range("E" & r), _
                               Address:=rawURL, _
                               TextToDisplay:="Link"
            ws.Range("E" & r).Font.Underline = xlUnderlineStyleSingle
            ws.Range("E" & r).Font.Color = RGB(34, 52, 86)   ' navy #223456
        Else
            ws.Range("E" & r).Font.Underline = xlUnderlineStyleNone
            ws.Range("E" & r).Font.Color = RGB(37, 37, 37)    ' #252525 body text
        End If

        ' 2) Mirror column D's FILL + BOTTOM BORDER onto E as the LAST step.
        '    Excel's built-in "Hyperlink" style (auto-stamped by Hyperlinks.Add)
        '    both whitens the fill AND strips the border - that's the "going
        '    white + gridlines removing" bug. D is coloured by the theme and is
        '    never touched here, so copying its band+border back guarantees E
        '    keeps the navy/gold band whether a link is present or not (incl.
        '    after a Reset). Done last so it always wins over the style.
        ws.Range("E" & r).Interior.Color = ws.Range("D" & r).Interior.Color
        ws.Range("E" & r).Font.Name = "Segoe UI"
        ws.Range("E" & r).Font.Size = 9
        With ws.Range("E" & r).Borders(xlEdgeBottom)
            .LineStyle = ws.Range("D" & r).Borders(xlEdgeBottom).LineStyle
            If ws.Range("D" & r).Borders(xlEdgeBottom).LineStyle <> xlNone Then
                .Weight = ws.Range("D" & r).Borders(xlEdgeBottom).Weight
                .Color = ws.Range("D" & r).Borders(xlEdgeBottom).Color
            End If
        End With
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RefreshSearchMatrixHyperlinks failed: " & Err.Description, vbCritical

End Sub


'====================================================================
'  REQUIRED FOR AUTO-REFRESH - paste this into SHEET1's own code
'  module (double-click "Sheet1" under Microsoft Excel Objects in the
'  VBA Project Explorer - NOT a standard module, and NOT this one).
'
'  It watches the specific input cells the Search Matrix depends on
'  (J9:J23 and T18:T23 - the customer/counterparty name, address,
'  and case-numbering fields) and rebuilds the links automatically
'  the instant any of them change, with no need to run the macro
'  yourself.
'====================================================================
'Private Sub Worksheet_Change(ByVal Target As Range)
'    Dim watched As Range
'    Set watched = Me.Range("J9:J23,T18:T23")
'
'    If Not Intersect(Target, watched) Is Nothing Then
'        RefreshSearchMatrixHyperlinks
'    End If
'End Sub


