Option Explicit

' ============================================================
' Run_Mass_Rename_v4
' Replacement launcher for the Rename button.
' File picking + validation is unchanged from Run_Mass_Rename_v3
' (Module7); everything after the picker now hands off to
' UserForm1 instead of the per-file InputBox wizard.
' Run_Mass_Rename_v3 / cleanFileName in Module7 are left in
' place untouched, and cleanFileName is reused as-is from there.
' ============================================================

Public Sub Run_Mass_Rename_v4()

    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo CancelHandler

    Dim ws As Worksheet: Set ws = ActiveSheet

    Dim ecm As String: ecm = ws.Range("J10").Value
    Dim sheetAlertID As String: sheetAlertID = ws.Range("J11").Value
    Dim custName As String: custName = ws.Range("J14").Value

    If ecm = "" Or sheetAlertID = "" Then
        MsgBox "Fill ECM (J10) & Alert ID (J11)!", vbCritical
        Exit Sub
    End If

    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)

    On Error Resume Next ' same guard Module7 uses around the picker setup
    With fd
        .title = "Select files to rename (Excel, PDF, Images, etc.)"
        .Filters.Clear
        .Filters.Add "All Files", "*.*", 1
        .AllowMultiSelect = True
    End With
    On Error GoTo CancelHandler

    If fd.Show <> -1 Then GoTo CancelHandler ' picker cancelled - no side effects
    If fd.SelectedItems.Count = 0 Then GoTo CancelHandler

    ' Snapshot the counterparty names now so the form doesn't depend on
    ' the sheet/selection staying put while the user works through it.
    Dim cpNames(1 To 6) As String
    Dim i As Long
    For i = 1 To 6
        cpNames(i) = ws.Range("J19").Offset(i - 1, 0).Value
    Next i

    UserForm1.InitRows fd.SelectedItems, ecm, sheetAlertID, custName, cpNames
    UserForm1.Show vbModal

    Application.EnableCancelKey = xlInterrupt
    Exit Sub

CancelHandler:
    Application.EnableCancelKey = xlInterrupt
    If Err.Number = 18 Then
        MsgBox "Process safely cancelled." & vbCrLf & vbCrLf & "You pressed ESC to stop.", vbInformation, "Aborted"
    ElseIf Err.Number <> 0 Then
        MsgBox "An unexpected error occurred:" & vbCrLf & Err.Description, vbCritical, "Error " & Err.Number
    End If
End Sub
