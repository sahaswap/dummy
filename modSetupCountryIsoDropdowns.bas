Option Explicit

' ================================================================
' modSetupCountryIsoDropdowns - ONE-TIME dashboard setup (re-runnable).
'
' Makes ISO Code 2 (L27, merged L27:N27) and ISO Code 3 (O27) behave
' exactly like the Country Name cell (G27):
'   - same GREY fill (copied from G27, so it always matches)
'   - UNLOCKED, so you can type in them under sheet protection
'   - list dropdown: ISO2 <- P36:P289, ISO3 <- Q36:Q289
'   - old XLOOKUP formulas cleared (they become input cells)
'
' Risk Score (P27) / Risk Classes (R27,T27) keep their formulas off G27.
'
' Each step is guarded, so one hiccup can't abort the rest - if anything
' goes wrong it's listed at the end instead of a blank "failed". Safe to
' run more than once. Needs the updated Sheet3 code-behind for the
' three fields to drive each other. Delete this module when done.
' ================================================================
Sub SetupCountryIsoDropdowns()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Sheet1")   ' tab name (VBA code-name Sheet3)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Dashboard sheet 'Sheet1' not found.", vbCritical
        Exit Sub
    End If

    Dim report As String
    Application.EnableEvents = False
    On Error Resume Next

    Err.Clear: ThisWorkbook.Unprotect Password:="p7ss"
    Err.Clear: ws.Unprotect Password:="p7ss"
    If Err.Number <> 0 Then report = report & "- unprotect sheet: " & Err.Description & vbCrLf

    ' 1. Match Country Name's grey (resolved RGB, so it's an exact match).
    Err.Clear
    Dim grey As Long: grey = ws.Range("G27").Interior.Color
    ws.Range("L27").Interior.Color = grey
    ws.Range("O27").Interior.Color = grey
    If Err.Number <> 0 Then report = report & "- grey fill: " & Err.Description & vbCrLf

    ' 2. Drop the old XLOOKUP formulas (these cells become inputs).
    '    Clear the FULL merged range for L27 - clearing just the top-left
    '    of a merged cell raises "cannot change part of a merged cell".
    Err.Clear
    ws.Range("L27:N27").ClearContents
    ws.Range("O27").ClearContents
    If Err.Number <> 0 Then report = report & "- clear formulas: " & Err.Description & vbCrLf

    ' 3. Unlock so they're typeable under protection.
    Err.Clear
    ws.Range("L27:N27").Locked = False
    ws.Range("O27").Locked = False
    If Err.Number <> 0 Then report = report & "- unlock: " & Err.Description & vbCrLf

    ' 4. List dropdowns, configured the same way as G27's.
    Err.Clear: SetListValidation ws.Range("L27:N27"), "=$P$36:$P$289"
    If Err.Number <> 0 Then report = report & "- ISO2 dropdown: " & Err.Description & vbCrLf
    Err.Clear: SetListValidation ws.Range("O27"), "=$Q$36:$Q$289"
    If Err.Number <> 0 Then report = report & "- ISO3 dropdown: " & Err.Description & vbCrLf

    ' 5. Re-protect (ignore "already protected" quirks).
    Err.Clear: ws.Protect Password:="p7ss"
    Err.Clear: ThisWorkbook.Protect Password:="p7ss", Structure:=True

    On Error GoTo 0
    Application.EnableEvents = True

    If report = "" Then
        MsgBox "Done." & vbCrLf & vbCrLf & _
               "ISO Code 2 (L27) and ISO Code 3 (O27) are now grey, unlocked " & _
               "dropdown cells - type in them or pick from the list, and all " & _
               "three fields (Country / ISO2 / ISO3) sync together.", _
               vbInformation, "Setup Complete"
    Else
        MsgBox "Setup finished, but these steps reported an issue:" & vbCrLf & vbCrLf & _
               report & vbCrLf & "Tell me which line above and I'll fix it precisely.", _
               vbExclamation, "Setup Report"
    End If
End Sub

Private Sub SetListValidation(ByVal target As Range, ByVal src As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=src
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub
