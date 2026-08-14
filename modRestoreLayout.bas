Option Explicit
'=====================================================================
' modRestoreLayout - ONE-TIME fix for Sheet1 after the 3.5 reformatting.
'
' The reformatting shifted the whole data block up one row, so every
' macro was reading the WRONG cells (ECM read the Alert ID, Decision read
' the Template, Customer Name read the Address, counterparties off by one,
' etc.). This puts the data back where all the code expects it - WITHOUT
' disturbing your new colors / fonts / borders (Excel carries formatting,
' merges, formulas, dropdowns and buttons along with the rows).
'
' What it restores:
'   Decision   J5 -> J6      Template   J6 -> J7
'   ProgName   J7 -> J8      EscScen    J8 -> J9
'   ECM Case   J9 -> J10     Alert ID   J10-> J11
'   Cust Name  J13-> J14     Cust Addr  J14-> J15
'   CP names   J18:J23 -> J19:J24   CP addr T18:T23 -> T19:T24
'   Master tbl O37:O290 -> O36:O289   (Country / G27 unchanged)
'
' Mechanics (rows 25 & 32 are verified empty):
'   Rows(5).Insert  -> top block down 1 (country/master over-shift)
'   Rows(25).Delete -> country + master back up 1
'   Rows(32).Delete -> master up 1 more, to its original O36:O289
'
' >>> RUN THIS ON A COPY of the workbook FIRST and eyeball it. <<<
'=====================================================================
Sub RestoreSheet1Layout()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Sheet1")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Sheet1 not found.", vbCritical: Exit Sub

    ' Guard: only run on the shifted layout (Decision label sits in G5).
    ' After the fix G5 is blank, so this can't be applied twice.
    If UCase$(Trim$(CStr(ws.Range("G5").Value))) <> "DECISION" Then
        MsgBox "This doesn't look like the shifted layout " & _
               "(expected the 'Decision' label in G5)." & vbCrLf & vbCrLf & _
               "Nothing was changed. If it's already fixed, you're done; " & _
               "otherwise tell me what's in G5 and I'll adjust.", vbExclamation, "Restore Layout"
        Exit Sub
    End If

    Dim prevCalc As Long
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail
    ws.Unprotect Password:="p7ss"

    ws.Rows(5).Insert                 ' top block -> down 1
    ws.Rows(25).Delete                ' pull country + master back up 1
    ws.Rows(32).Delete                ' pull master up 1 more (-> O36:O289)

    ws.Protect Password:="p7ss"

    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CalculateFull

    MsgBox "Sheet1 layout restored." & vbCrLf & vbCrLf & _
           "Check that 'Decision' is now in G6 and your data lines up, then " & _
           "test a button (e.g. Start) to confirm the ECM/Alert/Customer read " & _
           "correctly. If anything looks off, use your copy.", vbInformation, "Restore Layout"
    Exit Sub

Fail:
    Dim d As String: d = Err.Description
    On Error Resume Next
    ws.Protect Password:="p7ss"
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox "Restore failed: " & d & vbCrLf & vbCrLf & _
           "The sheet may be partly changed - press Ctrl+Z to undo, or use your copy.", _
           vbCritical, "Restore Layout"
End Sub
