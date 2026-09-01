Attribute VB_Name = "modThemeManager"
Option Explicit
'=====================================================================
' modThemeManager - one active theme at a time, cleanly.
'
' WHY THIS EXISTS
' Each theme module saves the "original" look of every shape it touches
' (into that shape's AltText) and of every cell (into a hidden backup
' sheet), so it can put things back. That only works if the theme is
' applied to an UNTHEMED sheet.
'
' Apply Tron on top of a live Navy & Gold sheet and Tron records Navy &
' Gold's navy fills and gold fonts as "the original". Removing Tron then
' returns the sheet to Navy & Gold - and Navy & Gold's own stash has
' been overwritten, so it can never be removed either. The two themes
' are now permanently fused, which is exactly the state that produced
' gold corner folds sitting on a black Grid canvas.
'
' The fix is a single entry point. Every Apply here removes whatever is
' currently active FIRST, then applies the requested theme to a clean
' baseline. Themes coexist in the workbook; they never overlap on the
' sheet.
'
' WIRE THE BUTTONS TO THESE, not to the theme modules directly:
'     ApplyTron        ApplyNavyGold_Safe        ClearTheme
'
' EVERY cross-module call below goes through Application.Run, on purpose.
' Do not "tidy" them into qualified calls like modTronLegacy.ApplyTronLegacy.
' A qualified call is resolved at COMPILE time and needs a module of that
' exact name to exist: if the .bas was pasted into the editor instead of
' imported (File > Import File), the module is called Module15 or similar,
' the qualifier resolves to nothing, and Option Explicit reports it as
' "Variable not defined" - a compile error that stops the whole project,
' not just this call. Application.Run resolves by PROCEDURE name at run
' time, so it works whatever the module is called, and it also lets this
' manager tolerate a theme module that is not installed at all.
'=====================================================================
Public Const THEME_NONE As String = "NONE"
Public Const THEME_TRON As String = "TRON"
Public Const THEME_NAVY As String = "NAVYGOLD"

Private Const STATE_NAME As String = "_ActiveTheme"

'---- public entry points --------------------------------------------
Sub ApplyTron()
    If Not ClearTheme(True) Then Exit Sub
    On Error GoTo Fail
    Application.Run "ApplyTronLegacy"
    SetActiveTheme THEME_TRON
    Exit Sub
Fail:
    MsgBox "Tron Legacy failed to apply: " & Err.Description, vbCritical, "Theme Manager"
End Sub

Sub ApplyNavyGold_Safe()
    If Not ClearTheme(True) Then Exit Sub
    On Error GoTo Fail
    Application.Run "ApplyNavyGold"
    SetActiveTheme THEME_NAVY
    Exit Sub
Fail:
    MsgBox "Navy & Gold failed to apply: " & Err.Description, vbCritical, "Theme Manager"
End Sub

' Strips whatever is active and leaves the sheet in its baseline look.
Sub RemoveActiveTheme()
    If ClearTheme(True) Then
        MsgBox "Theme removed. Sheet1 is back to its baseline look.", _
               vbInformation, "Theme Manager"
    End If
End Sub

' Returns False only if a removal genuinely failed, so callers can stop
' rather than stacking a second theme on a half-stripped sheet.
Public Function ClearTheme(ByVal quiet As Boolean) As Boolean
    Dim active As String
    active = ActiveTheme()
    ClearTheme = True

    On Error Resume Next
    Err.Clear

    ' Remove the recorded theme first, then sweep the others. A sweep is
    ' cheap - each Remove is a no-op when its own stash tag is absent -
    ' and it recovers a workbook whose state marker was lost or was
    ' themed before this manager existed.
    Select Case active
        Case THEME_TRON:  Application.Run "RemoveTronLegacy"
        Case THEME_NAVY:  Application.Run "RemoveNavyGold"
    End Select

    If active <> THEME_TRON Then Application.Run "RemoveTronLegacy"
    If active <> THEME_NAVY Then Application.Run "RemoveNavyGold"
    Application.Run "RemoveGlassStyle"          ' harmless if the module is absent

    UnhideDashboardShapes

    Err.Clear
    On Error GoTo 0

    SetActiveTheme THEME_NONE
End Function

' Repairs a shape that a theme hid and never un-hid.
'
' modNavyGold does not restyle the Beta title - RepositionBeta hides the
' workbook's real title shape and AddTitleBeta draws a gold replacement
' as NGADD_* shapes. RemoveNavyGold then deletes its replacement without
' ever setting the original back to visible, so removing Navy & Gold
' leaves the nav bar with no title at all. That is a bug in the theme
' module, but the manager is what promises a clean baseline, so the
' repair belongs here too rather than only inside one theme.
'
' Generated shapes are skipped by name prefix: those are meant to come
' and go with their theme.
Private Sub UnhideDashboardShapes()
    On Error Resume Next
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Sheets("Sheet1")
    If ws Is Nothing Then Exit Sub
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) <> "TRONADD_" And Left$(shp.Name, 6) <> "NGADD_" Then
            If shp.Visible = msoFalse Then shp.Visible = msoTrue
        End If
    Next shp
    On Error GoTo 0
End Sub

'---- state ----------------------------------------------------------
' Stored as a hidden workbook-level defined name: it survives save/close,
' never appears on a sheet, and needs no extra backup tab.
Public Function ActiveTheme() As String
    Dim v As String
    ActiveTheme = THEME_NONE
    On Error Resume Next
    v = ThisWorkbook.Names(STATE_NAME).RefersTo      ' arrives as ="TRON"
    On Error GoTo 0
    v = Replace(Replace(Replace(v, "=", ""), Chr$(34), ""), " ", "")
    If Len(v) > 0 Then ActiveTheme = UCase$(v)
End Function

Public Sub SetActiveTheme(ByVal v As String)
    On Error Resume Next
    ThisWorkbook.Unprotect Password:="p7ss"
    ThisWorkbook.Names(STATE_NAME).Delete
    ThisWorkbook.Names.Add Name:=STATE_NAME, RefersTo:="=""" & v & """", Visible:=False
    On Error GoTo 0
End Sub
