Option Explicit

' ================================================================
' modSetupCountryIsoDropdowns - ONE-TIME dashboard setup.
'
' Turns ISO Code 2 (L27, merged L27:N27) and ISO Code 3 (O27) into
' searchable list-dropdowns, matching the Country Name (G27) dropdown:
'   - ISO Code 2 list  <- P36:P289
'   - ISO Code 3 list  <- Q36:Q289
' Their old XLOOKUP formulas are cleared (they become input cells) and
' unlocked so they're editable under sheet protection. The Country<->ISO
' two-way sync in Sheet3's Worksheet_Change keeps all three in step.
'
' Risk Score (P27), PDF Risk Class (R27) and Excel Risk Class (T27) keep
' their formulas off G27 - untouched here.
'
' Run SetupCountryIsoDropdowns once, then this module can be deleted.
' NOTE: needs the updated Sheet3 code-behind in place for the dropdowns
' to actually drive each other.
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

    Application.EnableEvents = False
    On Error GoTo CleanFail
    ThisWorkbook.Unprotect Password:="p7ss"
    ws.Unprotect Password:="p7ss"

    ' ISO Code 2 -> list of column P ; ISO Code 3 -> list of column Q.
    SetListValidation ws.Range("L27"), "=$P$36:$P$289"
    SetListValidation ws.Range("O27"), "=$Q$36:$Q$289"

    ' They were XLOOKUP formulas; make them editable input cells instead.
    ws.Range("L27").ClearContents
    ws.Range("O27").ClearContents
    ws.Range("L27:N27").Locked = False    ' merged ISO2 cell
    ws.Range("O27").Locked = False        ' ISO3 cell

    ws.Protect Password:="p7ss"
    ThisWorkbook.Protect Password:="p7ss", Structure:=True
    Application.EnableEvents = True

    MsgBox "ISO Code 2 (L27) and ISO Code 3 (O27) are now dropdowns." & vbCrLf & vbCrLf & _
           "Pick a Country, an ISO2 or an ISO3 - the other two fill in, " & _
           "and Risk Score / Risk Classes follow automatically.", _
           vbInformation, "Setup Complete"
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    On Error Resume Next
    ws.Protect Password:="p7ss"
    ThisWorkbook.Protect Password:="p7ss", Structure:=True
    On Error GoTo 0
    MsgBox "Setup failed: " & Err.Description, vbCritical
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
