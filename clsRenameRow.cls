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
' WithEvents so a double-click re-opens the custom-name prompt for editing.
Public WithEvents LblPreview As MSForms.Label
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

' ---- CUSTOM NAMING (works for Word / PDF / images AND Excel) ----
' CAT_CUSTOM sits at the END of every entity's category list ->
'   ECM_AlertID_<Entity>_<user text>
' ENTITY_GENERAL is an extra entity at the end of the entity list; it has
' no identifier, so its only category is CAT_CUSTOM ->  ECM_AlertID_<user text>
' XL_CUSTOM is the Excel-type equivalent (Excel has no entity) ->
'   ECM_AlertID_<user text>
' The ECM_AlertID prefix is always kept (mandatory). The user-typed part is
' captured via a prompt and stored in m_CustomText; it is sanitised by
' cleanFileName in ComputeNewName just like every other generated name.
Private Const CAT_CUSTOM As String = "Custom Name..."
Private Const ENTITY_GENERAL As String = "General (Custom - no entity)"
Private Const XL_CUSTOM As String = "Custom Name..."

Private m_CustomText As String        ' the user-typed part after ECM_AlertID_
Private m_suppressPrompt As Boolean   ' True during CopyFrom so it never prompts

' Exposed so CopyFrom can read another row's custom text. VBA will NOT let
' you reach a Private member through an object reference (source.m_CustomText
' raises "Method or data member not found"), so route it through this property.
Public Property Get CustomText() As String
    CustomText = m_CustomText
End Property
Public Property Let CustomText(ByVal v As String)
    m_CustomText = v
End Property

' --- Called once by UserForm2 after the Excel-row controls are made:
'     fills the type dropdown, defaults to Alerted, sets field visibility.
Public Sub BuildXlTypeList()
    CboXlType.Clear
    CboXlType.AddItem XL_ALERTED
    CboXlType.AddItem XL_NONALERTED
    CboXlType.AddItem XL_LOOKBACK
    CboXlType.AddItem XL_GALILEO
    CboXlType.AddItem XL_CTS_KEEP
    CboXlType.AddItem XL_CUSTOM       ' -> ECM_AlertID_<custom text>
    CboXlType.ListIndex = 0            ' default: Alerted
    ApplyXlTypeVisibility
End Sub

' When a complete start date (mm.dd.yyyy) is entered, pre-populate the
' end date with the last day of the month one year later - e.g.
' 06.01.2025 -> 06.30.2026. Parsed by hand (Split on ".") so it doesn't
' depend on the machine's locale reading dotted dates. The analyst can
' still overwrite the end afterwards; it only re-derives if they edit
' the start again.
Private Sub AutoFillLookbackEnd()
    If TxtDateStart Is Nothing Or TxtDateEnd Is Nothing Then Exit Sub

    Dim parts() As String
    parts = Split(Trim(TxtDateStart.text), ".")
    If UBound(parts) <> 2 Then Exit Sub
    If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Sub

    Dim mm As Integer, dd As Integer, yyyy As Integer
    mm = CInt(parts(0)): dd = CInt(parts(1)): yyyy = CInt(parts(2))
    If yyyy < 1900 Or mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then Exit Sub

    On Error GoTo Done
    Dim startDate As Date, oneYr As Date, endDate As Date
    startDate = DateSerial(yyyy, mm, dd)
    oneYr = DateAdd("yyyy", 1, startDate)
    ' Day 0 of the following month = last day of oneYr's month.
    endDate = DateSerial(Year(oneYr), Month(oneYr) + 1, 0)
    TxtDateEnd.text = Format$(endDate, "mm.dd.yyyy")
Done:
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
    ' General custom entity (no identifier) -> ECM_AlertID_<custom text>
    CboEntity.AddItem ENTITY_GENERAL
    CboEntity.ListIndex = -1
End Sub

' True when the "General (Custom - no entity)" entity is the current choice.
Private Function IsGeneralEntity() As Boolean
    If CboEntity Is Nothing Then Exit Function
    IsGeneralEntity = (CboEntity.Value = ENTITY_GENERAL)
End Function

' --- Rebuilds the Category dropdown to match the current Entity choice ---
Private Sub RebuildCategoryList()
    Dim opts() As String
    Dim opt As Variant

    CboCategory.Clear
    If CboEntity.ListIndex < 0 Then Exit Sub

    ' General custom entity: the only category is the custom one, and we
    ' auto-select it so the prompt fires immediately (no extra click).
    If IsGeneralEntity() Then
        CboCategory.AddItem CAT_CUSTOM
        CboCategory.ListIndex = 0
        Exit Sub
    End If

    If CboEntity.ListIndex = 0 Then
        opts = Split(CUSTOMER_CATS, "|")
    Else
        opts = Split(CP_CATS, "|")
    End If

    For Each opt In opts
        CboCategory.AddItem opt
    Next opt
    ' every entity gets a Custom option at the END of its list
    CboCategory.AddItem CAT_CUSTOM
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
        ElseIf xlt = XL_CUSTOM Then
            ' Custom - ECM_AlertID prefix mandatory, rest is user-typed.
            If Trim(m_CustomText) = "" Then
                ComputeNewName = ""
                Exit Function
            End If
            newName = ECM & "_" & AlertID & "_" & m_CustomText
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
        If IsGeneralEntity() Then
            ' General custom - ECM_AlertID prefix only, no entity identifier.
            If Trim(m_CustomText) = "" Then
                ComputeNewName = ""
                Exit Function
            End If
            newName = ECM & "_" & AlertID & "_" & m_CustomText
        ElseIf CboCategory.Value = CAT_CUSTOM Then
            ' Per-entity custom - ECM_AlertID_<Entity>_<user text>.
            If Trim(m_CustomText) = "" Then
                ComputeNewName = ""
                Exit Function
            End If
            newName = ECM & "_" & AlertID & "_" & EntityIdentifier() & "_" & m_CustomText
        ElseIf CboCategory.Value = CAT_CTS_KYC Then
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
        If IsCustomActive() And Trim(m_CustomText) = "" Then
            LblPreview.caption = "(double-click to type a custom name)"
        ElseIf IsExcelType Then
            LblPreview.caption = "(enter both dates)"
        Else
            LblPreview.caption = "(select entity & category)"
        End If
        LblPreview.ForeColor = &H808080
    Else
        LblPreview.caption = result
        LblPreview.ForeColor = &H0
    End If
End Sub

' True when a Custom naming option is the current selection (Excel or not).
' Guards ListIndex < 0 first: reading .Value on an unselected combo returns
' Null, and assigning Null to this Boolean result would raise error 94.
Private Function IsCustomActive() As Boolean
    If IsExcelType Then
        If CboXlType Is Nothing Then Exit Function
        If CboXlType.ListIndex < 0 Then Exit Function
        IsCustomActive = (CboXlType.Value = XL_CUSTOM)
    Else
        If IsGeneralEntity() Then IsCustomActive = True: Exit Function
        If CboCategory Is Nothing Then Exit Function
        If CboCategory.ListIndex < 0 Then Exit Function
        IsCustomActive = (CboCategory.Value = CAT_CUSTOM)
    End If
End Function

' Prompt for the user-typed part. The ECM_AlertID prefix (and entity, when
' one applies) is shown so the analyst sees the full resulting name. Cancel
' keeps the previous value; a StrPtr of 0 marks Cancel vs an empty OK.
Private Sub PromptCustom()
    Dim prefix As String, s As String
    prefix = ECM & "_" & AlertID & "_"
    If (Not IsExcelType) And (Not IsGeneralEntity()) Then
        If CboEntity.ListIndex >= 0 Then prefix = prefix & EntityIdentifier() & "_"
    End If

    s = InputBox("Type the custom part of the file name." & vbCrLf & vbCrLf & _
                 "Saved as:  " & prefix & "<your text>" & FileExt, _
                 "Custom Name", m_CustomText)
    If StrPtr(s) = 0 Then Exit Sub        ' Cancelled - keep previous text
    m_CustomText = Trim(s)
End Sub

' --- Sets the outcome text/color after a rename attempt ---
Public Sub SetStatus(ByVal text As String, ByVal isError As Boolean)
    LblStatus.caption = text
    If isError Then
        LblStatus.ForeColor = &HFF&        ' red
    Else
        LblStatus.ForeColor = &H8000&      ' dark green
    End If
End Sub

' --- Copy-down: pull this row's Entity/Category or Sub-Alert from another row of the same type ---
Public Sub CopyFrom(ByVal source As clsRenameRow)
    If source.IsExcelType <> Me.IsExcelType Then Exit Sub ' only copy within same file type

    ' Suppress the custom prompt while we set combo values (their Change events
    ' would otherwise pop an InputBox on every row during "Apply to All").
    m_suppressPrompt = True
    m_CustomText = source.CustomText     ' via the public property - VBA blocks
                                         ' direct access to a Private member

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
    m_suppressPrompt = False
    RecalcPreview
End Sub

' ===================== Event handlers =====================

Private Sub CboEntity_Change()
    RebuildCategoryList
    RecalcPreview
End Sub

Private Sub CboCategory_Change()
    ' Selecting the per-entity "Custom Name..." (or the general custom, which
    ' is auto-selected) opens the prompt - unless we're mid-CopyFrom.
    If (CboCategory.Value = CAT_CUSTOM) And (Not m_suppressPrompt) Then PromptCustom
    RecalcPreview
End Sub

' Double-click the preview to re-open the prompt and edit the custom text.
Private Sub LblPreview_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsCustomActive() Then
        PromptCustom
        RecalcPreview
    End If
End Sub

Private Sub TxtSubAlert_Change()
    RecalcPreview
End Sub

Private Sub CboXlType_Change()
    ApplyXlTypeVisibility
    If (CboXlType.Value = XL_CUSTOM) And (Not m_suppressPrompt) Then PromptCustom
    RecalcPreview
End Sub

Private Sub TxtDateStart_Change()
    AutoFillLookbackEnd
    RecalcPreview
End Sub

' Double-click the start field to pick from the pop-up calendar. Setting
' the text fires TxtDateStart_Change, which auto-fills the end date and
' refreshes the preview. Guarded so that if frmCalendar isn't installed,
' the field just stays a normal typed textbox.
Private Sub TxtDateStart_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim d As Date
    If frmCalendar.PickDate(TxtDateStart.text, d) Then
        TxtDateStart.text = Format$(d, "mm.dd.yyyy")
    End If
End Sub

Private Sub TxtDateEnd_Change()
    RecalcPreview
End Sub
