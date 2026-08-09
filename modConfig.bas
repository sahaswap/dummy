'==================================================================
' modConfig  -  single source of truth for shared paths.
'
' The shared team tracker used to be hard-coded in FIVE modules
' (Module2, Module3, Module9, modRFIGenerate, modAuditLog). Every
' time it moved, all five had to be hunted down and edited - easy to
' miss one and silently send an analyst's data to the old location.
'
' Now those modules all call the functions below, so the tracker
' location lives in ONE place. If it ever moves again, change the
' single line in TrackerFolder() (and TrackerFileName() if the file
' is renamed) - nothing else needs touching.
'
' Functions (not Consts) because the path uses Environ("USERPROFILE"),
' which VBA does not allow inside a Const.
'==================================================================
Option Explicit

' The shared team folder (a locally-synced SharePoint library) that
' holds the tracker workbook and the per-analyst Audit_Logs mirror.
' >>> CHANGE THIS ONE LINE if the folder ever moves. <<<
Public Function TrackerFolder() As String
    TrackerFolder = Environ$("USERPROFILE") & _
        "\Community Federal Savings Bank\Mohini Srivastava - L1 Beta"
End Function

' The tracker workbook's file name - used both for the full path and
' for the "is it already open in this Excel session?" checks.
Public Function TrackerFileName() As String
    TrackerFileName = "Beta 2.4_Feedbacks & Issues Encountered.xlsx"
End Function

' Full path to the shared tracker workbook.
Public Function TrackerFile() As String
    TrackerFile = TrackerFolder() & "\" & TrackerFileName()
End Function
