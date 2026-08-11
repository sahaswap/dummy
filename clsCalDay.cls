Option Explicit

' ============================================================
' clsCalDay
' One instance per day cell in frmCalendar's month grid. Wraps
' the dynamically-created CommandButton so its click can be
' handled (MSForms has no control arrays, so each cell needs its
' own WithEvents sink - same trick clsRenameRow uses for rows).
' Kept alive in frmCalendar's m_cells collection.
' ============================================================

Public WithEvents Btn As MSForms.CommandButton
Public Parent As Object        ' the frmCalendar that owns this cell
Public TheDate As Date

Private Sub Btn_Click()
    If Not Parent Is Nothing Then Parent.SelectDate TheDate
End Sub
