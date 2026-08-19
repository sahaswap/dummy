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

' Short, plain-English guidance for picking a mode. Rewrites the verbose
' per-button description labels IN PLACE (matched by keyword in their
' current text) with crisp one-liners, and puts a short CAPTCHA tip next
' to the Cancel button. No form-designer work needed.
Private Sub UserForm_Initialize()
    On Error Resume Next

    ' Remove the old bottom guidance block, if a previous version added it.
    Me.Controls.Remove "lblGuide"

    ' Replace each verbose description label with a crisp one-liner. The
    ' existing labels are identified by a distinctive word in their text.
    Dim c As MSForms.Control, t As String
    For Each c In Me.Controls
        If TypeName(c) = "Label" Then
            t = LCase$(CStr(c.Caption))
            If InStr(t, "quickest") > 0 Then
                c.Caption = "Fastest - best on a good connection."
            ElseIf InStr(t, "bursty") > 0 Then
                c.Caption = "Steadier when the internet or VDI is slow / laggy."
            ElseIf InStr(t, "visible browser") > 0 Then
                c.Caption = "Most reliable (slowest) - non-English names, or when others keep getting blocked."
            End If
        End If
    Next c

    ' Short CAPTCHA tip, placed to the RIGHT of the Cancel button.
    Dim tip As MSForms.Label
    Set tip = Me.Controls.Add("Forms.Label.1", "lblGuide", True)
    tip.Left = btnCancel.Left + btnCancel.Width + 12
    tip.Top = btnCancel.Top
    tip.Width = Me.InsideWidth - tip.Left - 8
    tip.Height = 44
    tip.WordWrap = True
    tip.Font.Size = 8
    tip.Caption = "Tip: few CAPTCHAs & good internet -> use Fast (recommended). " & _
                  "Slow internet -> use Optimised (Fast on a poor connection triggers more CAPTCHAs)."

    ' tooltips too (harmless, on hover)
    btnFast.ControlTipText = "Fastest - best on a good connection."
    btnOptimised.ControlTipText = "Steadier when the internet or VDI is slow or laggy."
    btnVisible.ControlTipText = "Most reliable (slowest) - non-English names or when others keep getting blocked."

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

