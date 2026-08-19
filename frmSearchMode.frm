Attribute VB_Name = "frmSearchMode"
Attribute VB_Base = "0{6326C76B-E108-4C41-8A81-CA5783DF6D44}{B384715A-813C-4012-B8A6-9B0E712FF539}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Search mode chooser (v3.6+). Same button-choice pattern as
' UserForm1 (Transaction Files / Non Alerted / Cancel) - three
' named buttons instead of chained Yes/No/Cancel MsgBoxes.
'
' SelectedMode: 0 = Fast, 1 = Optimised, 2 = Visible.
' Only meaningful when userCancelled = False.
Public SelectedMode As Long
Public userCancelled As Boolean

' Numbered "Run Search 1/2/3" buttons (instead of Fast/Optimised/Visible)
' so analysts don't need to know the internals - they just try 1, then 2,
' then 3. Button 1 = Fast, 2 = Optimised, 3 = Visible under the hood; only
' the visible caption + guidance change. Verbose description labels are
' rewritten in place (matched by a word in their original text).
Private Sub UserForm_Initialize()
    On Error Resume Next

    ' Remove the old bottom guidance block, if a previous version added it.
    Me.Controls.Remove "lblGuide"

    ' Rename the buttons to simple numbered choices.
    btnFast.Caption = "Search 1"
    btnOptimised.Caption = "Search 2"
    btnVisible.Caption = "Search 3"

    ' Reword the title to make clear it's ONE choice (not three steps to run),
    ' and replace each verbose description label with a purely situational line
    ' (no "fast/steadier/reliable" wording - that would defeat the numbering).
    Dim c As MSForms.Control, t As String
    For Each c In Me.Controls
        If TypeName(c) = "Label" Then
            t = LCase$(CStr(c.Caption))
            If InStr(t, "choose a search") > 0 Or InStr(t, "search mode for this run") > 0 Then
                c.Caption = "Pick ONE search to run - start with Search 1:"
            ElseIf InStr(t, "quickest") > 0 Then
                c.Caption = "Start here - use this for most searches."
            ElseIf InStr(t, "bursty") > 0 Then
                c.Caption = "Use if Search 1 keeps getting blocked, or the internet / VDI is slow."
            ElseIf InStr(t, "visible browser") > 0 Then
                c.Caption = "Use for non-English names, or when the VDI is running its slowest."
            End If
        End If
    Next c

    ' Short ladder tip next to the Cancel button.
    Dim tip As MSForms.Label
    Set tip = Me.Controls.Add("Forms.Label.1", "lblGuide", True)
    tip.Left = btnCancel.Left + btnCancel.Width + 12
    tip.Top = btnCancel.Top
    tip.Width = Me.InsideWidth - tip.Left - 8
    tip.Height = 44
    tip.WordWrap = True
    tip.Font.Size = 8
    tip.Caption = "Try Search 1 first. If it keeps getting blocked or the connection is slow, move to Search 2, then Search 3."

    ' tooltips too (harmless, on hover)
    btnFast.ControlTipText = "Start here - use this for most searches."
    btnOptimised.ControlTipText = "Use if Search 1 keeps getting blocked, or the internet / VDI is slow."
    btnVisible.ControlTipText = "Use for non-English names, or when the VDI is running its slowest."

    On Error GoTo 0
End Sub

Private Sub btnFast_Click()
SelectedMode = 0
userCancelled = False
Me.Hide
End Sub

Private Sub btnOptimised_Click()
SelectedMode = 1
userCancelled = False
Me.Hide
End Sub

Private Sub btnVisible_Click()
SelectedMode = 2
userCancelled = False
Me.Hide
End Sub

Private Sub btnCancel_Click()
userCancelled = True
Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then userCancelled = True
End Sub

