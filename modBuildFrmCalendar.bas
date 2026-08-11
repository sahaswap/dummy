'==================================================================
' modBuildFrmCalendar - ONE-TIME setup macro. Run BuildFrmCalendar
' once to programmatically build the frmCalendar UserForm (prev/next
' buttons, month label, weekday headers, day-grid frame + full
' code-behind) without hand-building it in the VBA IDE.
'
' REQUIRES: File > Options > Trust Center > Trust Center Settings >
' Macro Settings > "Trust access to the VBA project object model"
' (same setting the other form builders need). If it errors
' immediately with "Could not access the VBA project", that setting
' isn't available and frmCalendar must be built manually.
'
' ALSO import clsCalDay (the day-cell click wrapper) - the grid won't
' respond to clicks without it.
'
' Safe to re-run: clears and rebuilds frmCalendar from scratch.
' Delete this module once frmCalendar is built.
'==================================================================
Option Explicit

Sub BuildFrmCalendar()
On Error GoTo TrustFail
Dim vbProj As Object
Set vbProj = Application.VBE.ActiveVBProject
On Error GoTo 0

Dim vbComp As Object
On Error Resume Next
Set vbComp = vbProj.VBComponents("frmCalendar")
On Error GoTo 0

If vbComp Is Nothing Then
Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm
vbComp.Name = "frmCalendar"
End If

vbComp.Properties("Width").Value = 236
vbComp.Properties("Height").Value = 214
vbComp.Properties("Caption").Value = "Pick a date"

Dim frm As Object
Set frm = vbComp.Designer

' Clear any existing controls so re-runs start clean.
Dim i As Long
For i = frm.Controls.Count - 1 To 0 Step -1
frm.Controls.Remove frm.Controls(i).Name
Next i

Dim btnPrev As Object, btnNext As Object, lblMonth As Object
Set btnPrev = frm.Controls.Add("Forms.CommandButton.1", "btnPrev")
btnPrev.Caption = "<": btnPrev.Left = 6: btnPrev.Top = 6: btnPrev.Width = 28: btnPrev.Height = 18

Set lblMonth = frm.Controls.Add("Forms.Label.1", "lblMonth")
lblMonth.Left = 38: lblMonth.Top = 8: lblMonth.Width = 156: lblMonth.Height = 16
lblMonth.TextAlign = 2 ' fmTextAlignCenter

Set btnNext = frm.Controls.Add("Forms.CommandButton.1", "btnNext")
btnNext.Caption = ">": btnNext.Left = 200: btnNext.Top = 6: btnNext.Width = 28: btnNext.Height = 18

' Weekday header row (Sun..Sat), aligned to the day columns below.
Dim dows As Variant, c As Long, hdr As Object
dows = Array("Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")
For c = 0 To 6
Set hdr = frm.Controls.Add("Forms.Label.1", "hdr" & c)
hdr.Left = 8 + c * 30: hdr.Top = 30: hdr.Width = 30: hdr.Height = 12
hdr.Caption = dows(c)
hdr.TextAlign = 2 ' center
Next c

' Frame that holds the day-number buttons (built at runtime).
Dim fraDays As Object
Set fraDays = frm.Controls.Add("Forms.Frame.1", "fraDays")
fraDays.Left = 6: fraDays.Top = 44: fraDays.Width = 214: fraDays.Height = 116
fraDays.Caption = ""

' ---- code-behind ----
Dim codeMod As Object
Set codeMod = vbComp.CodeModule
If codeMod.CountOfLines > 0 Then codeMod.DeleteLines 1, codeMod.CountOfLines

Dim s As String
s = ""
s = s & "Option Explicit" & vbCrLf
s = s & "" & vbCrLf
s = s & "' Lightweight month calendar. Shown modally via PickDate; returns" & vbCrLf
s = s & "' the chosen date. Day cells are built at runtime inside fraDays," & vbCrLf
s = s & "' each wrapped by a clsCalDay that reports its click back here. No" & vbCrLf
s = s & "' external ActiveX/OCX controls, so it works on locked-down machines." & vbCrLf
s = s & "" & vbCrLf
s = s & "Private m_view As Date       ' first day of the month currently shown" & vbCrLf
s = s & "Private m_picked As Boolean" & vbCrLf
s = s & "Private m_result As Date" & vbCrLf
s = s & "Private m_cells As Collection   ' clsCalDay instances (kept alive)" & vbCrLf
s = s & "" & vbCrLf
s = s & "' Show modally seeded from seedText (mm.dd.yyyy, else today)." & vbCrLf
s = s & "' Returns True + outDate if a day was picked, False if cancelled." & vbCrLf
s = s & "Public Function PickDate(ByVal seedText As String, ByRef outDate As Date) As Boolean" & vbCrLf
s = s & "Dim seed As Date" & vbCrLf
s = s & "seed = SeedFromText(seedText)" & vbCrLf
s = s & "m_view = DateSerial(Year(seed), Month(seed), 1)" & vbCrLf
s = s & "m_picked = False" & vbCrLf
s = s & "BuildGrid" & vbCrLf
s = s & "Me.Show 1" & vbCrLf
s = s & "PickDate = m_picked" & vbCrLf
s = s & "If m_picked Then outDate = m_result" & vbCrLf
s = s & "End Function" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Function SeedFromText(ByVal s As String) As Date" & vbCrLf
s = s & "Dim parts() As String" & vbCrLf
s = s & "parts = Split(Trim(s), ""."")" & vbCrLf
s = s & "If UBound(parts) = 2 Then" & vbCrLf
s = s & "If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then" & vbCrLf
s = s & "On Error Resume Next" & vbCrLf
s = s & "SeedFromText = DateSerial(CInt(parts(2)), CInt(parts(0)), CInt(parts(1)))" & vbCrLf
s = s & "If Err.Number = 0 Then Exit Function" & vbCrLf
s = s & "On Error GoTo 0" & vbCrLf
s = s & "End If" & vbCrLf
s = s & "End If" & vbCrLf
s = s & "SeedFromText = Date" & vbCrLf
s = s & "End Function" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub BuildGrid()" & vbCrLf
s = s & "lblMonth.Caption = Format$(m_view, ""mmmm yyyy"")" & vbCrLf
s = s & "' Drop the old cell wrappers first so their WithEvents refs are" & vbCrLf
s = s & "' released before we remove the buttons they point at." & vbCrLf
s = s & "Set m_cells = New Collection" & vbCrLf
s = s & "Dim i As Long" & vbCrLf
s = s & "For i = fraDays.Controls.Count - 1 To 0 Step -1" & vbCrLf
s = s & "fraDays.Controls.Remove fraDays.Controls(i).Name" & vbCrLf
s = s & "Next i" & vbCrLf
s = s & "" & vbCrLf
s = s & "Dim firstDow As Long, daysIn As Long, d As Long" & vbCrLf
s = s & "firstDow = Weekday(m_view, vbSunday) - 1" & vbCrLf
s = s & "daysIn = Day(DateSerial(Year(m_view), Month(m_view) + 1, 0))" & vbCrLf
s = s & "" & vbCrLf
s = s & "Const CW As Single = 30" & vbCrLf
s = s & "Const CH As Single = 18" & vbCrLf
s = s & "For d = 1 To daysIn" & vbCrLf
s = s & "Dim idx As Long, col As Long, row As Long" & vbCrLf
s = s & "idx = firstDow + (d - 1)" & vbCrLf
s = s & "col = idx Mod 7" & vbCrLf
s = s & "row = idx \ 7" & vbCrLf
s = s & "Dim btn As MSForms.CommandButton" & vbCrLf
s = s & "Set btn = fraDays.Controls.Add(""Forms.CommandButton.1"")" & vbCrLf
s = s & "btn.Left = col * CW" & vbCrLf
s = s & "btn.Top = row * CH" & vbCrLf
s = s & "btn.Width = CW" & vbCrLf
s = s & "btn.Height = CH" & vbCrLf
s = s & "btn.Caption = CStr(d)" & vbCrLf
s = s & "Dim cd As clsCalDay" & vbCrLf
s = s & "Set cd = New clsCalDay" & vbCrLf
s = s & "Set cd.Btn = btn" & vbCrLf
s = s & "Set cd.Parent = Me" & vbCrLf
s = s & "cd.TheDate = DateSerial(Year(m_view), Month(m_view), d)" & vbCrLf
s = s & "m_cells.Add cd" & vbCrLf
s = s & "Next d" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub btnPrev_Click()" & vbCrLf
s = s & "m_view = DateSerial(Year(m_view), Month(m_view) - 1, 1)" & vbCrLf
s = s & "BuildGrid" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub btnNext_Click()" & vbCrLf
s = s & "m_view = DateSerial(Year(m_view), Month(m_view) + 1, 1)" & vbCrLf
s = s & "BuildGrid" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "' Called by a clsCalDay when its day button is clicked." & vbCrLf
s = s & "Public Sub SelectDate(ByVal d As Date)" & vbCrLf
s = s & "m_result = d" & vbCrLf
s = s & "m_picked = True" & vbCrLf
s = s & "Me.Hide" & vbCrLf
s = s & "End Sub" & vbCrLf
s = s & "" & vbCrLf
s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
s = s & "If CloseMode = 0 Then       ' the X - cancel but keep the instance" & vbCrLf
s = s & "m_picked = False" & vbCrLf
s = s & "Cancel = True" & vbCrLf
s = s & "Me.Hide" & vbCrLf
s = s & "End If" & vbCrLf
s = s & "End Sub" & vbCrLf
codeMod.AddFromString s

MsgBox "frmCalendar built: prev/next + month label + weekday headers + day grid." & vbCrLf & vbCrLf & _
"Make sure clsCalDay is also imported, then you can delete this " & _
"modBuildFrmCalendar module - it's done its job.", vbInformation, "Build Complete"
Exit Sub

TrustFail:
MsgBox "Could not access the VBA project object model." & vbCrLf & vbCrLf & _
"Enable 'Trust access to the VBA project object model' in:" & vbCrLf & _
"File > Options > Trust Center > Trust Center Settings > Macro Settings." & vbCrLf & vbCrLf & _
"If that's locked by IT policy, frmCalendar must be built manually.", vbCritical, "Trust Setting Required"
End Sub
