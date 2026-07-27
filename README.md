Option Explicit

' ============================================================
' UserForm1 code-behind
' Builds every control at runtime from the empty form shell -
' no designer work, no MSComctl/MSFlexGrid OCX dependency.
' One clsRenameRow per selected file; each row owns its own
' controls and its own live-preview logic.
' ============================================================

Private Const COL_FILENAME_L As Long = 8
Private Const COL_FILENAME_W As Long = 150
Private Const COL_ENTITY_L As Long = 164
Private Const COL_ENTITY_W As Long = 130
Private Const COL_CATEGORY_L As Long = 300
Private Const COL_CATEGORY_W As Long = 150
Private Const COL_PREVIEW_L As Long = 456
Private Const COL_PREVIEW_W As Long = 230
Private Const COL_STATUS_L As Long = 692
Private Const COL_STATUS_W As Long = 90
Private Const ROW_H As Long = 20
Private Const HEADER_TOP As Long = 30
Private Const FRAME_TOP As Long = 50
Private Const FRAME_WIDTH As Long = 792
Private Const MAX_FRAME_HEIGHT As Long = 320

Private WithEvents btnCopyDown As MSForms.CommandButton
Private WithEvents btnRenameAll As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private rowsColl As Collection
Private frameRows As MSForms.Frame

' ============================================================
' Entry point called by Run_Mass_Rename_v4 before .Show
' ============================================================
Public Sub InitRows(ByVal selectedItems As Object, ByVal ecm As String, _
                     ByVal alertID As String, ByVal custName As String, _
                     ByRef cpNames() As String)

    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim item As Variant
    Dim idx As Long

    ClearForm

    Set rowsColl = New Collection
    Me.Caption = "Mass Rename - " & selectedItems.Count & " file(s) selected"

    BuildHeader
    BuildFrame
    BuildBottomButtons

    idx = 0
    For Each item In selectedItems
        idx = idx + 1
        AddRow idx, CStr(item), FSO, ecm, alertID, custName, cpNames
    Next item

    LayoutFrame idx
    SizeForm
End Sub

' --- Wipe everything so the form can be reused for a second run in the same session ---
Private Sub ClearForm()
    Dim names As Collection: Set names = New Collection
    Dim ctl As Control

    For Each ctl In Me.Controls
        names.Add ctl.Name
    Next ctl

    Dim nm As Variant
    For Each nm In names
        On Error Resume Next
        Me.Controls.Remove CStr(nm)
        On Error GoTo 0
    Next nm

    Set rowsColl = Nothing
    Set frameRows = Nothing
End Sub

Private Sub BuildHeader()
    Dim lbl As MSForms.Label

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblTitleHdr", True)
    With lbl
        .Left = 8: .Top = 8: .Width = 500: .Height = 16
        .Caption = "Review entity / category (or Sub-Alert ID) for each file, then Rename All."
        .Font.Bold = True
    End With

    AddHeaderLabel "lblHFN", "File", COL_FILENAME_L
    AddHeaderLabel "lblHEN", "Entity", COL_ENTITY_L
    AddHeaderLabel "lblHCT", "Category / Sub-Alert", COL_CATEGORY_L
    AddHeaderLabel "lblHPV", "New Name Preview", COL_PREVIEW_L
    AddHeaderLabel "lblHST", "Status", COL_STATUS_L
End Sub

Private Sub AddHeaderLabel(ByVal ctlName As String, ByVal text As String, ByVal leftPos As Long)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    With lbl
        .Left = leftPos: .Top = HEADER_TOP: .Width = 200: .Height = 14
        .Caption = text
        .Font.Bold = True
        .Font.Size = 8
    End With
End Sub

Private Sub BuildFrame()
    Set frameRows = Me.Controls.Add("Forms.Frame.1", "frameRows", True)
    With frameRows
        .Left = 6
        .Top = FRAME_TOP
        .Width = FRAME_WIDTH
        .Caption = ""
        .ScrollBars = fmScrollBarsVertical
        .SpecialEffect = fmSpecialEffectSunken
    End With
End Sub

Private Sub BuildBottomButtons()
    Set btnCopyDown = Me.Controls.Add("Forms.CommandButton.1", "btnCopyDown", True)
    With btnCopyDown
        .Caption = "Copy Row 1 Down to All"
        .Width = 150: .Height = 24: .Left = 8
    End With

    Set btnRenameAll = Me.Controls.Add("Forms.CommandButton.1", "btnRenameAll", True)
    With btnRenameAll
        .Caption = "Rename All"
        .Width = 100: .Height = 24
    End With

    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    With btnClose
        .Caption = "Cancel"
        .Width = 100: .Height = 24
    End With
End Sub

Private Sub AddRow(ByVal idx As Long, ByVal filePath As String, ByVal FSO As Object, _
                    ByVal ecm As String, ByVal alertID As String, ByVal custName As String, _
                    ByRef cpNames() As String)

    Dim r As New clsRenameRow
    Dim top As Long: top = (idx - 1) * ROW_H
    Dim ext As String: ext = "." & FSO.GetExtensionName(filePath)
    Dim fname As String: fname = FSO.GetFileName(filePath)
    Dim i As Long

    r.FilePath = filePath
    r.FileExt = ext
    r.IsExcelType = (LCase(ext) Like "*.xls*")
    r.ECM = ecm
    r.AlertID = alertID
    r.CustName = custName
    For i = 1 To 6
        r.CPNames(i) = cpNames(i)
    Next i

    Dim lbl As MSForms.Label
    Set lbl = frameRows.Controls.Add("Forms.Label.1", "lblFN" & idx, True)
    With lbl
        .Left = COL_FILENAME_L: .Top = top: .Width = COL_FILENAME_W: .Height = ROW_H - 2
        .Caption = fname
        .ControlTipText = filePath
        .WordWrap = False
        .AutoSize = False
    End With
    Set r.LblFileName = lbl

    If r.IsExcelType Then
        Dim txt As MSForms.TextBox
        Set txt = frameRows.Controls.Add("Forms.TextBox.1", "txtSA" & idx, True)
        With txt
            .Left = COL_ENTITY_L: .Top = top
            .Width = (COL_CATEGORY_L + COL_CATEGORY_W) - COL_ENTITY_L
            .Height = ROW_H - 2
            .ControlTipText = "Sub-Alert ID (optional)"
        End With
        Set r.TxtSubAlert = txt
    Else
        Dim cboE As MSForms.ComboBox
        Set cboE = frameRows.Controls.Add("Forms.ComboBox.1", "cboE" & idx, True)
        With cboE
            .Left = COL_ENTITY_L: .Top = top: .Width = COL_ENTITY_W: .Height = ROW_H - 2
            .Style = fmStyleDropDownList
        End With
        Set r.CboEntity = cboE
        r.BuildEntityList

        Dim cboC As MSForms.ComboBox
        Set cboC = frameRows.Controls.Add("Forms.ComboBox.1", "cboC" & idx, True)
        With cboC
            .Left = COL_CATEGORY_L: .Top = top: .Width = COL_CATEGORY_W: .Height = ROW_H - 2
            .Style = fmStyleDropDownList
        End With
        Set r.CboCategory = cboC
    End If

    Dim lblP As MSForms.Label
    Set lblP = frameRows.Controls.Add("Forms.Label.1", "lblPV" & idx, True)
    With lblP
        .Left = COL_PREVIEW_L: .Top = top: .Width = COL_PREVIEW_W: .Height = ROW_H - 2
        .WordWrap = False: .AutoSize = False
    End With
    Set r.LblPreview = lblP

    Dim lblS As MSForms.Label
    Set lblS = frameRows.Controls.Add("Forms.Label.1", "lblST" & idx, True)
    With lblS
        .Left = COL_STATUS_L: .Top = top: .Width = COL_STATUS_W: .Height = ROW_H - 2
        .WordWrap = False: .AutoSize = False
    End With
    Set r.LblStatus = lblS

    r.RecalcPreview

    rowsColl.Add r
End Sub

Private Sub LayoutFrame(ByVal rowCount As Long)
    Dim contentHeight As Long: contentHeight = rowCount * ROW_H
    Dim visibleHeight As Long: visibleHeight = contentHeight
    If visibleHeight > MAX_FRAME_HEIGHT Then visibleHeight = MAX_FRAME_HEIGHT
    If visibleHeight < ROW_H Then visibleHeight = ROW_H

    frameRows.Height = visibleHeight + 6
    frameRows.ScrollHeight = contentHeight + 4
    frameRows.ScrollWidth = FRAME_WIDTH - 20
End Sub

Private Sub SizeForm()
    Dim btnTop As Long
    btnTop = FRAME_TOP + frameRows.Height + 12

    btnCopyDown.Top = btnTop
    btnRenameAll.Top = btnTop
    btnClose.Top = btnTop

    btnRenameAll.Left = FRAME_WIDTH - 6 - btnClose.Width - 8 - btnRenameAll.Width
    btnClose.Left = FRAME_WIDTH - 6 - btnClose.Width

    Me.Width = FRAME_WIDTH + 24
    Me.Height = btnTop + 24 + 46
End Sub

' ============================================================
' Buttons
' ============================================================

Private Sub btnCopyDown_Click()
    If rowsColl Is Nothing Then Exit Sub
    If rowsColl.Count < 2 Then Exit Sub

    Dim firstRow As clsRenameRow: Set firstRow = rowsColl(1)
    Dim r As clsRenameRow
    Dim i As Long
    For i = 2 To rowsColl.Count
        Set r = rowsColl(i)
        r.CopyFrom firstRow
    Next i
End Sub

Private Sub btnRenameAll_Click()
    Dim r As clsRenameRow
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim newName As String, folder As String, targetPath As String, key As String

    ' Guard: don't run (or lock the form) if nothing is ready yet
    Dim anyReady As Boolean
    For Each r In rowsColl
        If r.ComputeNewName() <> "" Then anyReady = True: Exit For
    Next r
    If Not anyReady Then
        MsgBox "Select an Entity & Category (or leave the Sub-Alert ID as-is) for at least one file first.", vbExclamation
        Exit Sub
    End If

    ' Pass 1: compute targets, flag incompletes and in-batch duplicates
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    For Each r In rowsColl
        newName = r.ComputeNewName()
        If newName = "" Then
            r.SetStatus "Skipped: incomplete", True
        Else
            folder = FSO.GetParentFolderName(r.FilePath)
            targetPath = folder & "\" & newName
            key = LCase(targetPath)
            If Not dict.Exists(key) Then
                Dim grp As Collection: Set grp = New Collection
                dict.Add key, grp
            End If
            dict(key).Add r
        End If
    Next r

    Dim keyItem As Variant, rr As clsRenameRow
    For Each keyItem In dict.Keys
        If dict(keyItem).Count > 1 Then
            For Each rr In dict(keyItem)
                rr.SetStatus "Skipped: duplicate target", True
            Next rr
        End If
    Next keyItem

    ' Pass 2: rename everything that wasn't already flagged above
    For Each r In rowsColl
        If r.LblStatus.Caption <> "" Then GoTo ContinueLoop ' already flagged in pass 1

        newName = r.ComputeNewName()
        folder = FSO.GetParentFolderName(r.FilePath)
        targetPath = folder & "\" & newName

        If LCase(targetPath) = LCase(r.FilePath) Then
            r.SetStatus "No change needed", False
            GoTo ContinueLoop
        End If

        If FSO.FileExists(targetPath) Then
            r.SetStatus "Skipped: file exists", True
            GoTo ContinueLoop
        End If

        On Error Resume Next
        Err.Clear
        Name r.FilePath As targetPath
        If Err.Number = 0 Then
            r.SetStatus "Renamed", False
        Else
            r.SetStatus "ERROR: " & Err.Description, True
        End If
        On Error GoTo 0

ContinueLoop:
    Next r

    Dim renamedCount As Long, otherCount As Long
    For Each r In rowsColl
        If r.LblStatus.Caption = "Renamed" Then
            renamedCount = renamedCount + 1
        Else
            otherCount = otherCount + 1
        End If
    Next r

    MsgBox renamedCount & " file(s) renamed. " & otherCount & " not renamed - see the Status column.", _
           vbInformation, "Rename Complete"

    DisableAllControls
    btnClose.Caption = "Close"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub DisableAllControls()
    Dim r As clsRenameRow
    For Each r In rowsColl
        If r.IsExcelType Then
            r.TxtSubAlert.Enabled = False
        Else
            r.CboEntity.Enabled = False
            r.CboCategory.Enabled = False
        End If
    Next r
    btnCopyDown.Enabled = False
    btnRenameAll.Enabled = False
End Sub
