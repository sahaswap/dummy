Attribute VB_Name = "Module1"
Sub ClearForm()
Dim ws As Worksheet
Dim restoreErrNum As Long, restoreErrDesc As String

Set ws = ActiveSheet

' Mute Excel for speed during the clear.
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

On Error GoTo Restore

' UNLOCK AND CLEAR Sheet1 input fields only.
' ConsolidatedData is intentionally left intact - Sheet7's
' formulas pull from there.
ws.Unprotect Password:="p7ss"
' G27 = country; L27:N27 = ISO2 (full merge); O27 = ISO3. ISO2/ISO3 are
' input cells now (no longer XLOOKUP formulas), so they must be cleared
' explicitly - clearing the country alone no longer blanks them.
ws.Range("J6:J11, J14:J15, J19:J24, T19:T24, G27, L27:N27, O27").Value = ""
ws.Protect Password:="p7ss"
Restore:
restoreErrNum = Err.Number
restoreErrDesc = Err.Description

On Error Resume Next
' Best-effort re-protect in case the clear bailed mid-way.
ws.Protect Password:="p7ss"
On Error GoTo 0

' --- THE KEY CHANGE ---
' Force Automatic mode unconditionally. This corrects any leaked
' Manual-mode state from prior macros (OSINT, etc.) regardless
' of what the workbook came in with.
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True

On Error Resume Next
Application.CalculateFull
On Error GoTo 0

DoEvents

If restoreErrNum <> 0 Then
    MsgBox "ClearForm hit an error and stopped early:" & vbCrLf & vbCrLf & _
           "Err " & restoreErrNum & ": " & restoreErrDesc & vbCrLf & vbCrLf & _
           "Sheet1 may need to be re-protected manually.", _
           vbCritical, "ClearForm"
Else
    MsgBox "Form cleared.", vbInformation
End If
End Sub

