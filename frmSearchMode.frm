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

' Short, plain-English guidance for picking a mode. Shown BOTH as a
' hover-tooltip on each button AND as an always-visible label added at
' runtime below the buttons (so no form-designer work is needed). The
' label is positioned under whatever controls already exist, and the
' form grows to fit it - layout-agnostic, works however the buttons sit.
Private Sub UserForm_Initialize()
    On Error Resume Next

    ' tooltips (backup, on hover)
    btnFast.ControlTipText = "Fastest - your everyday default; best on a good connection."
    btnOptimised.ControlTipText = "Steadier when the internet or VDI is slow or laggy."
    btnVisible.ControlTipText = "Most reliable (but slowest) - non-English names or when others keep getting blocked."

    ' find the lowest existing control so the label sits just below it
    Dim c As MSForms.Control, maxBottom As Single
    maxBottom = 0
    For Each c In Me.Controls
        If (c.Top + c.Height) > maxBottom Then maxBottom = c.Top + c.Height
    Next c

    ' add the always-visible guidance label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblGuide", True)
    lbl.Left = 8
    lbl.Top = maxBottom + 8
    lbl.Width = Me.InsideWidth - 16
    lbl.Height = 84
    lbl.WordWrap = True
    lbl.Font.Size = 8
    lbl.Caption = _
        "Fast  -  fastest; best on a good connection." & vbCrLf & _
        "Optimised  -  steadier when the internet or VDI is slow / laggy." & vbCrLf & _
        "Visible  -  most reliable (slowest); non-English names or when others keep getting blocked." & vbCrLf & _
        "" & vbCrLf & _
        "Tip: few CAPTCHAs & good internet -> use Fast (recommended). Slow internet -> use Optimised. Running Fast on a poor connection triggers more CAPTCHAs."

    ' grow the form so the new label is fully visible
    Me.Height = Me.Height + lbl.Height + 16

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

