'==================================================================
' modBuildUserForm2  -  ONE-TIME setup macro. Run BuildUserForm2
' once to programmatically build UserForm2 (FrameRows + 3 buttons +
' full code-behind) without touching the VBA IDE's form designer by
' hand.
'
' REQUIRES: File > Options > Trust Center > Trust Center Settings >
' Macro Settings > check "Trust access to the VBA project object
' model". This is often locked by IT policy on managed machines -
' if BuildUserForm2 errors immediately with "Could not access the
' VBA project", that setting isn't available here and UserForm2
' needs to be built manually instead.
'
' Safe to re-run: clears and rebuilds UserForm2's controls and code
' from scratch each time, so re-running after any manual tweaks to
' the form will undo them.
'
' Delete this module once UserForm2 is built - it's a setup tool,
' not something the workbook needs going forward.
'==================================================================
Option Explicit

Sub BuildUserForm2()
On Error GoTo TrustFail
Dim vbProj As Object
Set vbProj = Application.VBE.ActiveVBProject
On Error GoTo 0

Dim vbComp As Object
On Error Resume Next
Set vbComp = vbProj.VBComponents("UserForm2")
On Error GoTo 0

If vbComp Is Nothing Then
Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm
vbComp.Name = "UserForm2"
End If

' Width/Height/Caption on a freshly-created form aren't reliably
' settable via .Designer.Width directly - the documented, robust
' way is through the component's Properties collection instead.
vbComp.Properties("Width").Value = 950
vbComp.Properties("Height").Value = 420
vbComp.Properties("Caption").Value = "Batch Rename"

Dim frm As Object
Set frm = vbComp.Designer

' Clear any existing controls so re-runs start from a clean slate.
Dim i As Long
For i = frm.Controls.Count - 1 To 0 Step -1
frm.Controls.Remove frm.Controls(i).Name
Next i

Dim fr As Object
Set fr = frm.Controls.Add("Forms.Frame.1", "FrameRows")
fr.Left = 6: fr.Top = 6: fr.Width = 930: fr.Height = 330
fr.ScrollBars = 2 ' fmScrollBarsVertical

Dim btn1 As Object, btn2 As Object, btn3 As Object
Set btn1 = frm.Controls.Add("Forms.CommandButton.1", "btnApplyToAll")
btn1.Caption = "Apply Row 1 to All"
btn1.Left = 6: btn1.Top = 345: btn1.Width = 140: btn1.Height = 26

Set btn2 = frm.Controls.Add("Forms.CommandButton.1", "btnRenameAll")
btn2.Caption = "Rename All"
btn2.Left = 155: btn2.Top = 345: btn2.Width = 100: btn2.Height = 26

Set btn3 = frm.Controls.Add("Forms.CommandButton.1", "btnCancel")
btn3.Caption = "Cancel"
btn3.Left = 265: btn3.Top = 345: btn3.Width = 100: btn3.Height = 26

' Replace the form's code-behind with the real host-form logic.
Dim codeMod As Object
Set codeMod = vbComp.CodeModule
If codeMod.CountOfLines > 0 Then codeMod.DeleteLines 1, codeMod.CountOfLines

Dim s As String
s = ""
s = s & "Option Explicit" & vbCrLf
s = s & "" & vbCrLf
s = s & "' ============================================================" & vbCrLf
s = s & "' UserForm2  -  batch rename host (Run_Mass_Rename_v4)" & vbCrLf
s = s & "'" & vbCrLf
s = s & "' Holds one clsRenameRow per file selected in Module7's" & vbCrLf
s = s & "' Run_Mass_Rename_v4 launcher. InitRows dynamically builds each" & vbCrLf
s = s & "' row's controls (Entity/Category combos for PDFs/images, an" & vbCrLf
s = s & "' Alerted/Lookback dropdown + Sub-Alert / date boxes for Excel" & vbCrLf
s = s & "' files, a live preview label, a" & vbCrLf
s = s & "' status label) inside FrameRows via Controls.Add - this is" & vbCrLf
s = s & "' ordinary MSForms runtime behavior, not project-structure" & vbCrLf
s = s & "' editing, so it doesn't need any special Trust Center setting." & vbCrLf
s = s & "'" & vbCrLf
s = s & "' Requires (built at design time in the VBA IDE's form designer -" & vbCrLf
s = s & "' see build steps given alongside this file):" & vbCrLf
s = s & "'   FrameRows      - a Frame control, ScrollBars = fmScrollBarsVertical" & vbCrLf
s = s & "'   btnApplyToAll  - CommandButton, ""Apply Row 1 to All""" & vbCrLf
s = s & "'   btnRenameAll   - CommandButton, ""Rename All""" & vbCrLf
s = s & "'   btnCancel      - CommandButton, ""Cancel""" & vbCrLf
s = s & "' ============================================================" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private rowsColl As Collection" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Const ROW_HEIGHT As Long = 22" & vbCrLf
s = s & "Private Const ROW_TOP_MARGIN As Long = 4" & vbCrLf
s = s & "" & vbCrLf
s = s & "' --- Called by Run_Mass_Rename_v4 before .Show ---" & vbCrLf
s = s & "Public Sub InitRows(ByVal selectedFiles As Object, ByVal ECM As String, _" & vbCrLf
s = s & "ByVal AlertID As String, ByVal CustName As String, CPNames() As String)" & vbCrLf
s = s & "Dim FSO As Object: Set FSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
s = s & "Set rowsColl = New Collection" & vbCrLf
s = s & "" & vbCrLf
s = s & "' Clear out anything left over from a previous run of this form." & vbCrLf
s = s & "Dim c As Long" & vbCrLf
s = s & "For c = FrameRows.Controls.Count - 1 To 0 Step -1" & vbCrLf
s = s & "FrameRows.Controls.Remove FrameRows.Controls(c).Name" & vbCrLf
s = s & "Next c" & vbCrLf
s = s & "" & vbCrLf
s = s & "Dim item As Variant, yPos As Long" & vbCrLf
s = s & "yPos = ROW_TOP_MARGIN" & vbCrLf
s = s & "" & vbCrLf
s = s & "For Each item In selectedFiles" & vbCrLf
s = s & "Dim r As clsRenameRow" & vbCrLf
s = s & "Set r = New clsRenameRow" & vbCrLf
s = s & "" & vbCrLf
s = s & "r.FilePath = CStr(item)" & vbCrLf
s = s & "r.FileExt = ""."" & FSO.GetExtensionName(CStr(item))" & vbCrLf
s = s & "r.IsExcelType = (LCase(r.FileExt) Like ""*.xls*"")" & vbCrLf
s = s & "r.ECM = ECM" & vbCrLf
s = s & "r.AlertID = AlertID" & vbCrLf
s = s & "r.CustName = CustName" & vbCrLf
s = s & "r.SetCPNames CPNames" & vbCrLf
s = s & "" & vbCrLf
s = s & "BuildRowControls r, FSO.GetFileName(CStr(item)), yPos" & vbCrLf
s = s & "If Not r.IsExcelType Then r.BuildEntityList" & vbCrLf
s = s & "r.RecalcPreview" & vbCrLf
s = s & "" & vbCrLf
s = s & "rowsColl.Add r" & vbCrLf
s = s & "yPos = yPos + ROW_HEIGHT" & vbCrLf
s = s & "Next item" & vbCrLf
s = s & "" & vbCrLf
s = s & "FrameRows.ScrollHeight = yPos + ROW_TOP_MARGIN" & vbCrLf
s = s & "Me.Caption = ""Batch Rename - "" & rowsColl.Count & "" file(s)""" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "' --- Dynamically adds this row's controls into FrameRows ---" & vbCrLf
s = s & "Private Sub BuildRowControls(ByRef r As clsRenameRow, ByVal displayName As String, ByVal yPos As Long)" & vbCrLf
s = s & "Dim lbl As MSForms.Label, lblPrev As MSForms.Label, lblStat As MSForms.Label" & vbCrLf
s = s & "" & vbCrLf
s = s & "Set lbl = FrameRows.Controls.Add(""Forms.Label.1"")" & vbCrLf
s = s & "lbl.Left = 6: lbl.Top = yPos: lbl.Width = 160: lbl.Height = 18" & vbCrLf
s = s & "lbl.Caption = displayName" & vbCrLf
s = s & "lbl.ControlTipText = displayName" & vbCrLf
s = s & "Set r.LblFileName = lbl" & vbCrLf
s = s & "" & vbCrLf
s = s & "If r.IsExcelType Then" & vbCrLf
s = s & "' Type dropdown (Alerted / Lookback) + the fields for each. Sub-Alert" & vbCrLf
s = s & "' and the two date boxes share the same zone; clsRenameRow shows only" & vbCrLf
s = s & "' the ones relevant to the selected type." & vbCrLf
s = s & "Dim cboX As MSForms.ComboBox" & vbCrLf
s = s & "Dim txtSA As MSForms.TextBox, txtD1 As MSForms.TextBox, txtD2 As MSForms.TextBox" & vbCrLf
s = s & "" & vbCrLf
s = s & "Set cboX = FrameRows.Controls.Add(""Forms.ComboBox.1"")" & vbCrLf
s = s & "cboX.Left = 170: cboX.Top = yPos: cboX.Width = 100: cboX.Height = 18" & vbCrLf
s = s & "cboX.Style = fmStyleDropDownList" & vbCrLf
s = s & "Set r.CboXlType = cboX" & vbCrLf
s = s & "" & vbCrLf
s = s & "Set txtSA = FrameRows.Controls.Add(""Forms.TextBox.1"")" & vbCrLf
s = s & "txtSA.Left = 275: txtSA.Top = yPos: txtSA.Width = 260: txtSA.Height = 18" & vbCrLf
s = s & "txtSA.ControlTipText = ""Sub-Alert ID (optional)""" & vbCrLf
s = s & "Set r.TxtSubAlert = txtSA" & vbCrLf
s = s & "" & vbCrLf
s = s & "Set txtD1 = FrameRows.Controls.Add(""Forms.TextBox.1"")" & vbCrLf
s = s & "txtD1.Left = 275: txtD1.Top = yPos: txtD1.Width = 125: txtD1.Height = 18" & vbCrLf
s = s & "txtD1.ControlTipText = ""Lookback START date (mm.dd.yyyy) - double-click for calendar""" & vbCrLf
s = s & "Set r.TxtDateStart = txtD1" & vbCrLf
s = s & "" & vbCrLf
s = s & "Set txtD2 = FrameRows.Controls.Add(""Forms.TextBox.1"")" & vbCrLf
s = s & "txtD2.Left = 405: txtD2.Top = yPos: txtD2.Width = 125: txtD2.Height = 18" & vbCrLf
s = s & "txtD2.ControlTipText = ""Lookback END date (mm.dd.yyyy)""" & vbCrLf
s = s & "Set r.TxtDateEnd = txtD2" & vbCrLf
s = s & "" & vbCrLf
s = s & "r.BuildXlTypeList" & vbCrLf
s = s & "Else" & vbCrLf
s = s & "Dim cboE As MSForms.ComboBox, cboC As MSForms.ComboBox" & vbCrLf
s = s & "Set cboE = FrameRows.Controls.Add(""Forms.ComboBox.1"")" & vbCrLf
s = s & "cboE.Left = 170: cboE.Top = yPos: cboE.Width = 180: cboE.Height = 18" & vbCrLf
s = s & "cboE.Style = fmStyleDropDownList" & vbCrLf
s = s & "Set r.CboEntity = cboE" & vbCrLf
s = s & "" & vbCrLf
s = s & "Set cboC = FrameRows.Controls.Add(""Forms.ComboBox.1"")" & vbCrLf
s = s & "cboC.Left = 355: cboC.Top = yPos: cboC.Width = 180: cboC.Height = 18" & vbCrLf
s = s & "cboC.Style = fmStyleDropDownList" & vbCrLf
s = s & "Set r.CboCategory = cboC" & vbCrLf
s = s & "End If" & vbCrLf
s = s & "" & vbCrLf
s = s & "Set lblPrev = FrameRows.Controls.Add(""Forms.Label.1"")" & vbCrLf
s = s & "lblPrev.Left = 540: lblPrev.Top = yPos: lblPrev.Width = 230: lblPrev.Height = 18" & vbCrLf
s = s & "Set r.LblPreview = lblPrev" & vbCrLf
s = s & "" & vbCrLf
s = s & "Set lblStat = FrameRows.Controls.Add(""Forms.Label.1"")" & vbCrLf
s = s & "lblStat.Left = 775: lblStat.Top = yPos: lblStat.Width = 140: lblStat.Height = 18" & vbCrLf
s = s & "Set r.LblStatus = lblStat" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "' --- Sets every row (after the first) to the first row's Entity/Category" & vbCrLf
s = s & "'     or Sub-Alert - CopyFrom() already skips mismatched file types. ---" & vbCrLf
s = s & "Private Sub btnApplyToAll_Click()" & vbCrLf
s = s & "If rowsColl Is Nothing Or rowsColl.Count < 2 Then Exit Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Dim first As clsRenameRow, r As clsRenameRow, i As Long" & vbCrLf
s = s & "Set first = rowsColl(1)" & vbCrLf
s = s & "" & vbCrLf
s = s & "For i = 2 To rowsColl.Count" & vbCrLf
s = s & "Set r = rowsColl(i)" & vbCrLf
s = s & "r.CopyFrom first" & vbCrLf
s = s & "Next i" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "' --- Renames every row that has a valid preview; leaves the rest alone" & vbCrLf
s = s & "'     with a status explaining why, instead of aborting the whole batch. ---" & vbCrLf
s = s & "Private Sub btnRenameAll_Click()" & vbCrLf
s = s & "If rowsColl Is Nothing Or rowsColl.Count = 0 Then Exit Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Dim FSO As Object: Set FSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
s = s & "Dim r As clsRenameRow, newName As String, successCount As Long" & vbCrLf
s = s & "" & vbCrLf
s = s & "For Each r In rowsColl" & vbCrLf
s = s & "newName = r.ComputeNewName()" & vbCrLf
s = s & "If newName = """" Then" & vbCrLf
s = s & "r.SetStatus ""Not ready - pick entity/category"", True" & vbCrLf
s = s & "Else" & vbCrLf
s = s & "On Error Resume Next" & vbCrLf
s = s & "Err.Clear" & vbCrLf
s = s & "Name r.FilePath As FSO.GetParentFolderName(r.FilePath) & ""\"" & newName" & vbCrLf
s = s & "If Err.Number = 0 Then" & vbCrLf
s = s & "r.SetStatus ""Renamed"", False" & vbCrLf
s = s & "successCount = successCount + 1" & vbCrLf
s = s & "Else" & vbCrLf
s = s & "r.SetStatus ""Failed: "" & Err.Description, True" & vbCrLf
s = s & "End If" & vbCrLf
s = s & "On Error GoTo 0" & vbCrLf
s = s & "End If" & vbCrLf
s = s & "Next r" & vbCrLf
s = s & "" & vbCrLf
s = s & "MsgBox successCount & "" of "" & rowsColl.Count & "" file(s) renamed."", _" & vbCrLf
s = s & "vbInformation, ""Rename Complete""" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub btnCancel_Click()" & vbCrLf
s = s & "Me.Hide" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
s = s & "' Nothing to persist on the form itself - files already renamed by" & vbCrLf
s = s & "' btnRenameAll are already done; this just closes the window." & vbCrLf
s = s & "End Sub" & vbCrLf
codeMod.AddFromString s

MsgBox "UserForm2 built: FrameRows + 3 buttons + code-behind all in place." & vbCrLf & vbCrLf & _
"You can now delete this modBuildUserForm2 module - it's done its job.", _
vbInformation, "Build Complete"
Exit Sub

TrustFail:
MsgBox "Could not access the VBA project object model." & vbCrLf & vbCrLf & _
"This needs 'Trust access to the VBA project object model' enabled in:" & vbCrLf & _
"File > Options > Trust Center > Trust Center Settings > Macro Settings." & vbCrLf & vbCrLf & _
"If that's locked by IT policy on this machine, UserForm2 needs to be built " & _
"manually instead - see the separate build-steps notes.", vbCritical, "Trust Setting Required"
End Sub
