Option Explicit

' ================================================================
' modFixCounterpartyLayout - one-time Sheet1 layout cleanup.
'
' Fixes the "Counterparty Information" section so it ends at column T:
'   1. The section header is a merged cell (top-left G17) that had been
'      merged all the way out to column AB. It is unmerged and re-merged
'      to G17:T17, matching the Alert / Customer headers above it.
'   2. The stray fill + borders that spilled into columns U:AB across the
'      counterparty block and the country-risk header row (the extra blue
'      cell at U26) are cleared, so nothing shows past column T.
'
' NON-DESTRUCTIVE: only merge state and formatting are changed - no cell
' CONTENTS are deleted. If any of the boundaries below are slightly off
' for your sheet, just tweak the constants and re-run.
'
' Run FixCounterpartyLayout once, then this module can be deleted.
' ================================================================
Private Const SHEET_NAME     As String = "Sheet1"
Private Const HEADER_TOPLEFT As String = "G17"       ' the merged header cell
Private Const HEADER_MERGE_TO As String = "G17:T17"  ' where the header should end
Private Const STRAY_RANGE    As String = "U17:AB27"  ' zone to strip past column T

Sub FixCounterpartyLayout()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & SHEET_NAME & "' not found.", vbCritical
        Exit Sub
    End If

    Dim caption As String
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1. Remember the header caption, then unmerge the over-wide header.
    caption = CStr(ws.Range(HEADER_TOPLEFT).Value)
    If caption = "" Then caption = "Counterparty Information"
    If ws.Range(HEADER_TOPLEFT).MergeCells Then _
        ws.Range(HEADER_TOPLEFT).MergeArea.UnMerge

    ' 2. Strip stray fill + borders past column T. Contents are left
    '    alone, so nothing is lost if a cell isn't actually empty.
    With ws.Range(STRAY_RANGE)
        .UnMerge
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With

    ' 3. Re-merge the header so it ends exactly at column T.
    ws.Range(HEADER_TOPLEFT).Value = caption
    ws.Range(HEADER_MERGE_TO).Merge
    With ws.Range(HEADER_TOPLEFT)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Done." & vbCrLf & vbCrLf & _
           "- 'Counterparty Information' header re-merged to end at column T." & vbCrLf & _
           "- Stray fill/borders in " & STRAY_RANGE & " cleared (incl. U26)." & vbCrLf & vbCrLf & _
           "If the right edge at column T now needs a border, or the address " & _
           "column actually sits beyond T, tell me and I'll adjust.", _
           vbInformation, "Layout Fixed"
End Sub
