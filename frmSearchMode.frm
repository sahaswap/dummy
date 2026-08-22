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

' Numbered "Option 1/2/3" buttons (instead of Fast/Optimised/Visible) so
' analysts don't need to know the internals. Option 1/2 positions are
' SWAPPED from the old Search 1/2 layout: Optimised is now first (Option
' 1), Fast is now second (Option 2). Button 3 = Visible under the hood;
' only the caption, guidance, and position change - SelectedMode values
' (0/1/2) are untouched, so ApplySearchMode's mapping in Module2 needs no
' changes. Guidance wording is deliberately situational only (which
' condition applies to you), never comparative ("better"/"steadier"/
' "quicker"/"default") - the analyst should pick based on their own
' situation, not be nudged toward one option. Verbose description labels
' are matched by a word in their ORIGINAL (Designer) text, since
' Initialize runs fresh every Show/Unload cycle.
Private Sub UserForm_Initialize()
    On Error Resume Next

    ' Remove the old bottom guidance block, if a previous version added it.
    Me.Controls.Remove "lblGuide"

    ' ---- uniform button size: the SMALLEST of the 4 current sizes, so
    ' nothing grows into a neighbouring label (only ever shrinks to match). ----
    Dim uW As Single, uH As Single
    uW = btnFast.Width: If btnOptimised.Width < uW Then uW = btnOptimised.Width
    If btnVisible.Width < uW Then uW = btnVisible.Width
    If btnCancel.Width < uW Then uW = btnCancel.Width
    uH = btnFast.Height: If btnOptimised.Height < uH Then uH = btnOptimised.Height
    If btnVisible.Height < uH Then uH = btnVisible.Height
    If btnCancel.Height < uH Then uH = btnCancel.Height
    btnFast.Width = uW: btnFast.Height = uH
    btnOptimised.Width = uW: btnOptimised.Height = uH
    btnVisible.Width = uW: btnVisible.Height = uH
    btnCancel.Width = uW: btnCancel.Height = uH

    ' ---- swap places: Optimised moves to where Fast was (now first),
    ' Fast moves to where Optimised was (now second). Visible is untouched. ----
    Dim tmpTop As Single
    tmpTop = btnFast.Top
    btnFast.Top = btnOptimised.Top
    btnOptimised.Top = tmpTop

    ' ---- enforce a UNIFORM gap between all 4 buttons. The swap above only
    ' exchanges Optimised/Fast's Top values - it never touched whatever gap
    ' the Designer originally had between Fast/Visible and Visible/Cancel,
    ' which is why 2-3 and 3-4 had a much bigger gap than 1-2. All four
    ' Tops are now recomputed from Optimised's (now topmost) position using
    ' one fixed gap, so every row is spaced identically. ----
    Const BTN_GAP As Single = 10
    Dim topAnchor As Single
    topAnchor = btnOptimised.Top
    btnFast.Top = topAnchor + (uH + BTN_GAP)
    btnVisible.Top = topAnchor + 2 * (uH + BTN_GAP)
    btnCancel.Top = topAnchor + 3 * (uH + BTN_GAP)

    ' Rename the buttons to match their NEW visual order (1 = whichever is
    ' now on top = Optimised, 2 = Fast, 3 = Visible, unchanged).
    btnOptimised.Caption = "Option 1"
    btnFast.Caption = "Option 2"
    btnVisible.Caption = "Option 3"

    ' Reword the title, and replace each verbose description label with a
    ' short situational line - matched by ORIGINAL wording, THEN moved and
    ' resized to sit centered against its (possibly just-swapped) button.
    Dim c As MSForms.Control, t As String
    Dim lblFast As MSForms.Label, lblOptimised As MSForms.Label, lblVisible As MSForms.Label
    For Each c In Me.Controls
        If TypeName(c) = "Label" Then
            t = LCase$(CStr(c.Caption))
            If InStr(t, "choose a search") > 0 Or InStr(t, "search mode for this run") > 0 Then
                c.Caption = "Pick ONE search to run, based on your situation:"
            ElseIf InStr(t, "quickest") > 0 Then
                Set lblFast = c   ' this label sits beside the Fast button
            ElseIf InStr(t, "bursty") > 0 Then
                Set lblOptimised = c   ' this label sits beside the Optimised button
            ElseIf InStr(t, "visible browser") > 0 Then
                Set lblVisible = c
                c.Caption = "Use for non-English names, or as a last resort."
            End If
        End If
    Next c

    ' Observable-EVENT wording, not environment-quality wording. "Slow/
    ' laggy" vs "good/fast" still carried an implicit good-vs-bad judgment
    ' about the analyst's connection - even without comparative words like
    ' "steadier"/"quicker", one condition read as a problem and the other
    ' as ideal. Tying the choice to something DIRECTLY OBSERVED (a CAPTCHA
    ' block, yes or no) is a neutral fact, not a quality judgment, so
    ' neither option reads as the "good" or "bad" one to be in.
    If Not lblOptimised Is Nothing Then
        lblOptimised.Caption = "Use if you are seeing repeated CAPTCHA blocks."
    End If
    If Not lblFast Is Nothing Then
        lblFast.Caption = "Use if you are not seeing CAPTCHA blocks."
    End If

    ' ---- move + center each description label against its OWN button's
    ' new position. A fixed guessed Height (the old approach) centers the
    ' BOX against the button, but MSForms.Label has no vertical-align
    ' property - text always starts at the TOP of whatever Height is set.
    ' If the box is taller than the actual wrapped text, the box looks
    ' centered but the readable text still sits high. CenterLabelOnButton
    ' fixes this by letting AutoSize measure the TRUE wrapped-text height
    ' at the label's fixed Width first, then centers THAT exact size. ----
    If Not lblOptimised Is Nothing Then CenterLabelOnButton lblOptimised, btnOptimised
    If Not lblFast Is Nothing Then CenterLabelOnButton lblFast, btnFast
    If Not lblVisible Is Nothing Then CenterLabelOnButton lblVisible, btnVisible

    ' Short usage tip next to the Cancel button - purely situational, no
    ' option named as the default or preferred choice. Styled like a
    ' terms-and-conditions footnote (small, italic, asterisk-prefixed)
    ' so it reads as fine-print reference, not a headline instruction.
    Dim tip As MSForms.Label
    Set tip = Me.Controls.Add("Forms.Label.1", "lblGuide", True)
    tip.Left = btnCancel.Left + btnCancel.Width + 12
    tip.Top = btnCancel.Top
    tip.Width = Me.InsideWidth - tip.Left - 8
    tip.Height = 44
    tip.WordWrap = True
    tip.Font.Size = 7
    tip.Font.Italic = True
    tip.Caption = "*Option 1 - you're seeing repeated CAPTCHA blocks. Option 2 - you're not seeing CAPTCHA blocks. Option 3 - non-English names, or as a last resort."

    ' tooltips too (harmless, on hover)
    btnOptimised.ControlTipText = "Use if you are seeing repeated CAPTCHA blocks."
    btnFast.ControlTipText = "Use if you are not seeing CAPTCHA blocks."
    btnVisible.ControlTipText = "Use for non-English names, or as a last resort."

    On Error GoTo 0
End Sub

' Vertically centers lbl against btn using lbl's TRUE rendered text height,
' not a guessed constant. AutoSize + WordWrap=True (with Width already
' fixed) makes the Forms engine shrink/grow Height to exactly fit the
' wrapped text at that Width - only Height changes, Width stays put. That
' accurate Height is then centered against the button's box, so the
' visible TEXT (not just an oversized bounding box) lines up with the
' button's middle.
Private Sub CenterLabelOnButton(ByVal lbl As MSForms.Label, ByVal btn As MSForms.CommandButton)
    On Error Resume Next
    lbl.TextAlign = 2 ' fmTextAlignCenter
    lbl.WordWrap = True
    lbl.AutoSize = True     ' measures the real wrapped-text height at the fixed Width
    lbl.AutoSize = False    ' lock that height so nothing can resize it again later
    lbl.Top = btn.Top + (btn.Height - lbl.Height) / 2
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

