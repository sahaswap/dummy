Option Explicit

' ============================================================
' clsRenameRow
' One instance per file selected in Run_Mass_Rename_v4.
' Owns that row's controls, knows how to rebuild its own
' Category list when Entity changes, and recomputes its own
' live preview. Kept alive in UserForm2's rowsColl so the
' WithEvents hooks below don't get garbage-collected.
' ============================================================

Public WithEvents CboEntity As MSForms.ComboBox
Public WithEvents CboCategory As MSForms.ComboBox
Public WithEvents TxtSubAlert As MSForms.TextBox

' Excel-row extras: a type dropdown (Alerted vs Lookback) and the two
' Lookback date fields, which are only shown when Lookback is chosen.
Public WithEvents CboXlType As MSForms.ComboBox
Public WithEvents TxtDateStart As MSForms.TextBox
Public WithEvents TxtDateEnd As MSForms.TextBox

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

' Object modules (class modules, UserForms, etc.) can't have a
' Public array as a member - VBA compile error "Constants,
' fixed-length strings, arrays, user-defined types and Declare
' statements not allowed as Public members of object modules".
' Kept Private instead, set via SetCPNames() from the launcher.
Private m_CPNames(1 To 6) As String

Private Const CUSTOMER_CATS As String = "External Search 1|External Search 2|External Search 3|Website|Google Translate|Customer Name + SSN|Sigma|Alert Write Up|Galileo Transaction"
Private Const CP_CATS As String = "External Search 1|External Search 2|External Search 3|Website|Google Translate"

' Excel type-dropdown options.
Private Const XL_ALERTED As String = "Alerted Transactions"
Private Const XL_LOOKBACK As String = "Lookback Transactions"

' --- Called once by UserForm2 after the Excel-row controls are made:
'     fills the type dropdown, defaults to Alerted, sets field visibility.
Public Sub BuildXlTypeList()
    CboXlType.Clear
    CboXlType.AddItem XL_ALERTED
    CboXlType.AddItem XL_LOOKBACK
    CboXlType.ListIndex = 0            ' default: Alerted
    ApplyXlTypeVisibility
End Sub

' Shows the Sub-Alert box for Alerted, the two date boxes for Lookback.
Private Sub ApplyXlTypeVisibility()
    Dim isLookback As Boolean
    isLookback = (CboXlType.Value = XL_LOOKBACK)
    If Not TxtSubAlert Is Nothing Then TxtSubAlert.Visible = Not isLookback
    If Not TxtDateStart Is Nothing Then TxtDateStart.Visible = isLookback
    If Not TxtDateEnd Is Nothing Then TxtDateEnd.Visible = isLookback
End Sub

' --- Called once by UserForm2's InitRows to copy in the 6 counterparty names ---
Public Sub SetCPNames(ByRef names() As String)
    Dim i As Long
    For i = 1 To 6
        m_CPNames(i) = names(i)
    Next i
End Sub

' --- Called once by UserForm1 right after the controls are created ---
Public Sub BuildEntityList()
    Dim i As Long
    CboEntity.Clear
    CboEntity.AddItem "Customer (" & CustName & ")"
    For i = 1 To 6
        If m_CPNames(i) <> "" Then
            CboEntity.AddItem "Counterparty " & i & " (" & m_CPNames(i) & ")"
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
        cpIdx = CboEntity.ListIndex ' 1-6, list positions line up with m_CPNames
        If m_CPNames(cpIdx) = "" Then
            EntityIdentifier = "Counterparty " & cpIdx
        Else
            EntityIdentifier = "Counterparty " & cpIdx & "_" & m_CPNames(cpIdx)
        End If
    End If
End Function

' --- Returns the final cleaned target file name (with extension), or "" if not ready ---
Public Function ComputeNewName() As String
    Dim newName As String

    If IsExcelType Then
        If Not CboXlType Is Nothing And CboXlType.Value = XL_LOOKBACK Then
            ' Lookback - need both dates before the name is valid.
            Dim ds As String, de As String
            ds = Trim(TxtDateStart.text)
            de = Trim(TxtDateEnd.text)
            If ds = "" Or de = "" Then
                ComputeNewName = ""
                Exit Function
            End If
            newName = ECM & "_" & AlertID & "_Lookback Transactions (" & _
                      FormatLookbackDate(ds) & " to " & FormatLookbackDate(de) & ")"
        Else
            ' Alerted (default) - Sub-Alert ID optional.
            Dim subAlert As String: subAlert = Trim(TxtSubAlert.text)
            If subAlert <> "" Then
                newName = ECM & "_" & AlertID & "_Alerted Transactions(" & subAlert & ")"
            Else
                newName = ECM & "_" & AlertID & "_Alerted Transactions"
            End If
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
    If LblPreview Is Nothing Then Exit Sub   ' controls not fully built yet
    Dim result As String
    result = ComputeNewName()
    If result = "" Then
        If IsExcelType Then
            LblPreview.Caption = "(enter both dates)"
        Else
            LblPreview.Caption = "(select entity & category)"
        End If
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
        If Not source.CboXlType Is Nothing And Not CboXlType Is Nothing Then
            CboXlType.Value = source.CboXlType.Value
        End If
        ApplyXlTypeVisibility
        TxtSubAlert.text = source.TxtSubAlert.text
        If Not source.TxtDateStart Is Nothing And Not TxtDateStart Is Nothing Then _
            TxtDateStart.text = source.TxtDateStart.text
        If Not source.TxtDateEnd Is Nothing And Not TxtDateEnd Is Nothing Then _
            TxtDateEnd.text = source.TxtDateEnd.text
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

Private Sub CboXlType_Change()
    ApplyXlTypeVisibility
    RecalcPreview
End Sub

Private Sub TxtDateStart_Change()
    RecalcPreview
End Sub

Private Sub TxtDateEnd_Change()
    RecalcPreview
End Sub
