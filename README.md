Option Explicit

' ============================================================
' clsRenameRow
' One instance per file selected in Run_Mass_Rename_v4.
' Owns that row's controls, knows how to rebuild its own
' Category list when Entity changes, and recomputes its own
' live preview. Kept alive in UserForm1's rowsColl so the
' WithEvents hooks below don't get garbage-collected.
' ============================================================

Public WithEvents CboEntity As MSForms.ComboBox
Public WithEvents CboCategory As MSForms.ComboBox
Public WithEvents TxtSubAlert As MSForms.TextBox

Public LblFileName As MSForms.Label
Public LblPreview As MSForms.Label
Public LblStatus As MSForms.Label

Public FilePath As String      ' full original path
Public FileExt As String       ' includes leading dot, e.g. ".pdf"
Public IsExcelType As Boolean

' Shared context, copied in from the launcher at InitRow time
Public ECM As String
Public AlertID As String
Public CustName As String
Public CPNames(1 To 6) As String

Private Const CUSTOMER_CATS As String = "External Search 1|External Search 2|External Search 3|Website|Google Translate|Customer Name + SSN|Sigma|Alert Write Up|Galileo Transaction"
Private Const CP_CATS As String = "External Search 1|External Search 2|External Search 3|Website|Google Translate"

' --- Called once by UserForm1 right after the controls are created ---
Public Sub BuildEntityList()
    Dim i As Long
    CboEntity.Clear
    CboEntity.AddItem "Customer (" & CustName & ")"
    For i = 1 To 6
        If CPNames(i) <> "" Then
            CboEntity.AddItem "Counterparty " & i & " (" & CPNames(i) & ")"
        Else
            CboEntity.AddItem "Counterparty " & i
        End If
    Next i
    CboEntity.ListIndex = -1
End Sub

' --- Rebuilds the Category dropdown to match the current Entity choice ---
Private Sub RebuildCategoryList()
    Dim opts() As String
    Dim opt As Variant

    CboCategory.Clear
    If CboEntity.ListIndex < 0 Then Exit Sub

    If CboEntity.ListIndex = 0 Then
        opts = Split(CUSTOMER_CATS, "|")
    Else
        opts = Split(CP_CATS, "|")
    End If

    For Each opt In opts
        CboCategory.AddItem opt
    Next opt
    CboCategory.ListIndex = -1
End Sub

' --- Builds the identifier text exactly the way Run_Mass_Rename_v3 does ---
Private Function EntityIdentifier() As String
    Dim cpIdx As Long
    If CboEntity.ListIndex = 0 Then
        EntityIdentifier = CustName
    ElseIf CboEntity.ListIndex >= 1 Then
        cpIdx = CboEntity.ListIndex ' 1-6, list positions line up with CPNames
        If CPNames(cpIdx) = "" Then
            EntityIdentifier = "Counterparty " & cpIdx
        Else
            EntityIdentifier = "Counterparty " & cpIdx & "_" & CPNames(cpIdx)
        End If
    End If
End Function

' --- Returns the final cleaned target file name (with extension), or "" if not ready ---
Public Function ComputeNewName() As String
    Dim newName As String

    If IsExcelType Then
        Dim subAlert As String: subAlert = Trim(TxtSubAlert.Text)
        If subAlert <> "" Then
            newName = ECM & "_" & AlertID & "_Alerted Transactions(" & subAlert & ")"
        Else
            newName = ECM & "_" & AlertID & "_Alerted Transactions"
        End If
    Else
        If CboEntity.ListIndex < 0 Or CboCategory.ListIndex < 0 Then
            ComputeNewName = ""
            Exit Function
        End If
        newName = ECM & "_" & AlertID & "_" & EntityIdentifier() & "_" & CboCategory.Value
    End If

    ComputeNewName = cleanFileName(newName) & FileExt
End Function

' --- Refreshes the on-screen preview label for this row ---
Public Sub RecalcPreview()
    Dim result As String
    result = ComputeNewName()
    If result = "" Then
        LblPreview.Caption = "(select entity & category)"
        LblPreview.ForeColor = &H808080
    Else
        LblPreview.Caption = result
        LblPreview.ForeColor = &H0
    End If
End Sub

' --- Sets the outcome text/color after a rename attempt ---
Public Sub SetStatus(ByVal text As String, ByVal isError As Boolean)
    LblStatus.Caption = text
    If isError Then
        LblStatus.ForeColor = &HFF&        ' red
    Else
        LblStatus.ForeColor = &H8000&      ' dark green
    End If
End Sub

' --- Copy-down: pull this row's Entity/Category or Sub-Alert from another row of the same type ---
Public Sub CopyFrom(ByVal source As clsRenameRow)
    If source.IsExcelType <> Me.IsExcelType Then Exit Sub ' only copy within same file type

    If IsExcelType Then
        TxtSubAlert.Text = source.TxtSubAlert.Text
    Else
        CboEntity.ListIndex = source.CboEntity.ListIndex
        RebuildCategoryList
        If source.CboCategory.ListIndex >= 0 And source.CboCategory.ListIndex < CboCategory.ListCount Then
            CboCategory.ListIndex = source.CboCategory.ListIndex
        End If
    End If
    RecalcPreview
End Sub

' ===================== Event handlers =====================

Private Sub CboEntity_Change()
    RebuildCategoryList
    RecalcPreview
End Sub

Private Sub CboCategory_Change()
    RecalcPreview
End Sub

Private Sub TxtSubAlert_Change()
    RecalcPreview
End Sub
