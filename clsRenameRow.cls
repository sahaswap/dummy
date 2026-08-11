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

' Galileo moved to the Excel type dropdown (it's a spreadsheet, not a
' PDF), and Sigma replaced by the CTS Report / Customer KYC label.
Private Const CUSTOMER_CATS As String = "External Search 1|External Search 2|External Search 3|Website|Google Translate|Customer Name + SSN|CTS Report_Customer KYC|Alert Write Up|Galileo Profile"
Private Const CP_CATS As String = "External Search 1|External Search 2|External Search 3|Website|Google Translate"

' This customer category is a standalone document label - it
' deliberately drops the entity identifier from the final file name,
' producing ECM_AlertID_CTS Report_Customer KYC.
Private Const CAT_CTS_KYC As String = "CTS Report_Customer KYC"

' Excel type-dropdown options.
Private Const XL_ALERTED As String = "Alerted Transactions"
Private Const XL_NONALERTED As String = "Non Alerted Transactions"
Private Const XL_LOOKBACK As String = "Lookback Transactions"
Private Const XL_GALILEO As String = "Galileo Transaction"
' Keeps the xlsx's own (static, system-generated) name and just
' prepends "ECM_AlertID_CTS Report_" in front of it, producing
' ECM_AlertID_CTS Report_<original name>. Label is picker-only text.
Private Const XL_CTS_KEEP As String = "CTS Report (keep file name)"

' --- Called once by UserForm2 after the Excel-row controls are made:
'     fills the type dropdown, defaults to Alerted, sets field visibility.
Public Sub BuildXlTypeList()
    CboXlType.Clear
    CboXlType.AddItem XL_ALERTED
    CboXlType.AddItem XL_NONALERTED
    CboXlType.AddItem XL_LOOKBACK
    CboXlType.AddItem XL_GALILEO
    CboXlType.AddItem XL_CTS_KEEP
    CboXlType.ListIndex = 0            ' default: Alerted
    ApplyXlTypeVisibility
End Sub

' Sub-Alert box only for Alerted, the two date boxes only for Lookback.
' Non Alerted and Galileo are fixed names - they need no extra input.
Private Sub ApplyXlTypeVisibility()
    Dim isLookback As Boolean, isAlerted As Boolean
    isLookback = (CboXlType.Value = XL_LOOKBACK)
    isAlerted = (CboXlType.Value = XL_ALERTED)
    If Not TxtSubAlert Is Nothing Then TxtSubAlert.Visible = isAlerted
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

' --- The file's own name minus path and extension (for prefix-style
'     renames that keep the original static name). ---
Private Function OriginalBaseName() As String
    Dim fn As String
    fn = Mid$(FilePath, InStrRev(FilePath, "\") + 1)   ' name + extension
    If Len(FileExt) > 0 And Len(fn) >= Len(FileExt) Then
        If LCase$(Right$(fn, Len(FileExt))) = LCase$(FileExt) Then
            fn = Left$(fn, Len(fn) - Len(FileExt))
        End If
    End If
    OriginalBaseName = fn
End Function

' --- Returns the final cleaned target file name (with extension), or "" if not ready ---
Public Function ComputeNewName() As String
    Dim newName As String

    If IsExcelType Then
        Dim xlt As String
        If Not CboXlType Is Nothing Then xlt = CboXlType.Value

        If xlt = XL_LOOKBACK Then
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
        ElseIf xlt = XL_NONALERTED Then
            newName = ECM & "_" & AlertID & "_Non Alerted Transactions"
        ElseIf xlt = XL_GALILEO Then
            newName = ECM & "_" & AlertID & "_Galileo Transaction"
        ElseIf xlt = XL_CTS_KEEP Then
            ' Keep the file's own static name, just prepend the prefix.
            newName = ECM & "_" & AlertID & "_CTS Report_" & OriginalBaseName()
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
        If CboCategory.Value = CAT_CTS_KYC Then
            ' Standalone label - no entity identifier in the file name.
            newName = ECM & "_" & AlertID & "_" & CboCategory.Value
        Else
            newName = ECM & "_" & AlertID & "_" & EntityIdentifier() & "_" & CboCategory.Value
        End If
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
