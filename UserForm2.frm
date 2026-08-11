Option Explicit

' ============================================================
' UserForm2  -  batch rename host (Run_Mass_Rename_v4)
'
' Holds one clsRenameRow per file selected in Module7's
' Run_Mass_Rename_v4 launcher. InitRows dynamically builds each
' row's controls (Entity/Category combos for PDFs/images, a
' Sub-Alert textbox for Excel files, a live preview label, a
' status label) inside FrameRows via Controls.Add - this is
' ordinary MSForms runtime behavior, not project-structure
' editing, so it doesn't need any special Trust Center setting.
'
' Requires (built at design time in the VBA IDE's form designer -
' see build steps given alongside this file):
'   FrameRows      - a Frame control, ScrollBars = fmScrollBarsVertical
'   btnApplyToAll  - CommandButton, "Apply Row 1 to All"
'   btnRenameAll   - CommandButton, "Rename All"
'   btnCancel      - CommandButton, "Cancel"
' ============================================================

Private rowsColl As Collection

Private Const ROW_HEIGHT As Long = 22
Private Const ROW_TOP_MARGIN As Long = 4

' --- Called by Run_Mass_Rename_v4 before .Show ---
Public Sub InitRows(ByVal selectedFiles As Object, ByVal ECM As String, _
ByVal AlertID As String, ByVal CustName As String, CPNames() As String)
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
Set rowsColl = New Collection

' Clear out anything left over from a previous run of this form.
Dim c As Long
For c = FrameRows.Controls.Count - 1 To 0 Step -1
FrameRows.Controls.Remove FrameRows.Controls(c).Name
Next c

Dim item As Variant, yPos As Long
yPos = ROW_TOP_MARGIN

For Each item In selectedFiles
Dim r As clsRenameRow
Set r = New clsRenameRow

r.FilePath = CStr(item)
r.FileExt = "." & FSO.GetExtensionName(CStr(item))
r.IsExcelType = (LCase(r.FileExt) Like "*.xls*")
r.ECM = ECM
r.AlertID = AlertID
r.CustName = CustName
r.SetCPNames CPNames

BuildRowControls r, FSO.GetFileName(CStr(item)), yPos
If Not r.IsExcelType Then r.BuildEntityList
r.RecalcPreview

rowsColl.Add r
yPos = yPos + ROW_HEIGHT
Next item

FrameRows.ScrollHeight = yPos + ROW_TOP_MARGIN
Me.Caption = "Batch Rename - " & rowsColl.Count & " file(s)"
End Sub

' --- Dynamically adds this row's controls into FrameRows ---
Private Sub BuildRowControls(ByRef r As clsRenameRow, ByVal displayName As String, ByVal yPos As Long)
Dim lbl As MSForms.Label, lblPrev As MSForms.Label, lblStat As MSForms.Label

Set lbl = FrameRows.Controls.Add("Forms.Label.1")
lbl.Left = 6: lbl.Top = yPos: lbl.Width = 160: lbl.Height = 18
lbl.Caption = displayName
lbl.ControlTipText = displayName
Set r.LblFileName = lbl

If r.IsExcelType Then
' Type dropdown (Alerted / Lookback) + the fields for each. Sub-Alert
' and the two date boxes share the same zone; clsRenameRow shows only
' the ones relevant to the selected type.
Dim cboX As MSForms.ComboBox
Dim txtSA As MSForms.TextBox, txtD1 As MSForms.TextBox, txtD2 As MSForms.TextBox

Set cboX = FrameRows.Controls.Add("Forms.ComboBox.1")
cboX.Left = 170: cboX.Top = yPos: cboX.Width = 100: cboX.Height = 18
cboX.Style = fmStyleDropDownList
Set r.CboXlType = cboX

Set txtSA = FrameRows.Controls.Add("Forms.TextBox.1")
txtSA.Left = 275: txtSA.Top = yPos: txtSA.Width = 260: txtSA.Height = 18
txtSA.ControlTipText = "Sub-Alert ID (optional)"
Set r.TxtSubAlert = txtSA

Set txtD1 = FrameRows.Controls.Add("Forms.TextBox.1")
txtD1.Left = 275: txtD1.Top = yPos: txtD1.Width = 125: txtD1.Height = 18
txtD1.ControlTipText = "Lookback START date (mm.dd.yyyy)"
Set r.TxtDateStart = txtD1

Set txtD2 = FrameRows.Controls.Add("Forms.TextBox.1")
txtD2.Left = 405: txtD2.Top = yPos: txtD2.Width = 125: txtD2.Height = 18
txtD2.ControlTipText = "Lookback END date (mm.dd.yyyy)"
Set r.TxtDateEnd = txtD2

r.BuildXlTypeList   ' fill dropdown, default Alerted, set field visibility
Else
Dim cboE As MSForms.ComboBox, cboC As MSForms.ComboBox
Set cboE = FrameRows.Controls.Add("Forms.ComboBox.1")
cboE.Left = 170: cboE.Top = yPos: cboE.Width = 180: cboE.Height = 18
cboE.Style = fmStyleDropDownList
Set r.CboEntity = cboE

Set cboC = FrameRows.Controls.Add("Forms.ComboBox.1")
cboC.Left = 355: cboC.Top = yPos: cboC.Width = 180: cboC.Height = 18
cboC.Style = fmStyleDropDownList
Set r.CboCategory = cboC
End If

Set lblPrev = FrameRows.Controls.Add("Forms.Label.1")
lblPrev.Left = 540: lblPrev.Top = yPos: lblPrev.Width = 230: lblPrev.Height = 18
Set r.LblPreview = lblPrev

Set lblStat = FrameRows.Controls.Add("Forms.Label.1")
lblStat.Left = 775: lblStat.Top = yPos: lblStat.Width = 140: lblStat.Height = 18
Set r.LblStatus = lblStat
End Sub

' --- Sets every row (after the first) to the first row's Entity/Category
'     or Sub-Alert - CopyFrom() already skips mismatched file types. ---
Private Sub btnApplyToAll_Click()
If rowsColl Is Nothing Or rowsColl.Count < 2 Then Exit Sub

Dim first As clsRenameRow, r As clsRenameRow, i As Long
Set first = rowsColl(1)

For i = 2 To rowsColl.Count
Set r = rowsColl(i)
r.CopyFrom first
Next i
End Sub

' --- Renames every row that has a valid preview; leaves the rest alone
'     with a status explaining why, instead of aborting the whole batch. ---
Private Sub btnRenameAll_Click()
If rowsColl Is Nothing Or rowsColl.Count = 0 Then Exit Sub

Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
Dim r As clsRenameRow, newName As String, successCount As Long

For Each r In rowsColl
newName = r.ComputeNewName()
If newName = "" Then
r.SetStatus "Not ready - pick entity/category", True
Else
On Error Resume Next
Err.Clear
Name r.FilePath As FSO.GetParentFolderName(r.FilePath) & "\" & newName
If Err.Number = 0 Then
r.SetStatus "Renamed", False
successCount = successCount + 1
Else
r.SetStatus "Failed: " & Err.Description, True
End If
On Error GoTo 0
End If
Next r

MsgBox successCount & " of " & rowsColl.Count & " file(s) renamed.", _
vbInformation, "Rename Complete"
End Sub

Private Sub btnCancel_Click()
Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' Nothing to persist on the form itself - files already renamed by
' btnRenameAll are already done; this just closes the window.
End Sub
