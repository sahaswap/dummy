Option Explicit

' ================================================================
' modTrackerPush - shared "append a row to the OneDrive master tracker"
' logic, plus DEFERRED (Application.OnTime) wrappers so the button macros
' return immediately and the tracker updates a moment later in the
' background instead of blocking the analyst on a slow network write.
'
'   PushRFITracker_Deferred  -> tracker Sheet3  (Generate Narrative)
'   PushTrxTracker_Deferred  -> tracker Sheet2  (Export Trx File)
'
' Both read the ECM ID fresh from Sheet1!J10 (it won't change in the ~1s
' between scheduling and firing). PushTrackerRow does the actual write
' via a hidden "ghost" Excel instance, exactly as the old inline code did.
' Failures are swallowed - a tracker hiccup must never disrupt the
' analyst, and this now runs detached from the main flow.
' ================================================================

Public Sub PushRFITracker_Deferred()
    PushTrackerRow "Sheet3", TrackerEcmID()
End Sub

Public Sub PushTrxTracker_Deferred()
    PushTrackerRow "Sheet2", TrackerEcmID()
End Sub

Private Function TrackerEcmID() As String
    On Error Resume Next
    TrackerEcmID = Trim(ThisWorkbook.Sheets("Sheet1").Range("J10").Value)
    On Error GoTo 0
End Function

' Appends one row (Now | USERNAME | ecmID | version) to the tracker's
' targetSheet. Reuses an already-open copy if present, else opens the
' tracker in THIS Excel instance with its window hidden. Opening it in a
' SEPARATE ("ghost") Excel is what caused the "Microsoft Excel is waiting
' for another application to complete an OLE action" dialog when the
' OneDrive file was slow to open - there is no cross-process OLE here.
Public Sub PushTrackerRow(ByVal targetSheet As String, ByVal ecmID As String)
    Dim masterWb As Workbook, masterWs As Worksheet, pushWb As Workbook
    Dim mRow As Long, wasAlreadyOpen As Boolean, openedHidden As Boolean
    Dim expectedHeaders As Variant, hdrIdx As Integer, headersOK As Boolean, headerMsg As String
    Dim prevEvents As Boolean, prevScreen As Boolean

    On Error GoTo CleanExit
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False       ' also suppresses the tracker's Workbook_Open
    Application.ScreenUpdating = False

    wasAlreadyOpen = False
    For Each pushWb In Application.Workbooks
        If StrComp(pushWb.Name, TrackerFileName(), vbTextCompare) = 0 Then
            Set masterWb = pushWb
            wasAlreadyOpen = True
            Exit For
        End If
    Next pushWb

    If masterWb Is Nothing Then
        On Error Resume Next
        Set masterWb = Application.Workbooks.Open(fileName:=TrackerFile(), UpdateLinks:=False)
        On Error GoTo CleanExit
        If masterWb Is Nothing Then GoTo CleanExit
        openedHidden = True
        On Error Resume Next
        masterWb.Windows(1).Visible = False   ' keep it out of sight
        On Error GoTo CleanExit
    End If

    If masterWb.ReadOnly Then
        If openedHidden Then masterWb.Close SaveChanges:=False
        GoTo CleanExit
    End If

    On Error Resume Next
    Set masterWs = masterWb.Sheets(targetSheet)
    On Error GoTo CleanExit
    If masterWs Is Nothing Then
        If openedHidden Then masterWb.Close SaveChanges:=False
        GoTo CleanExit
    End If

    expectedHeaders = Array("Date & Time", "Analyst ID", "ECM Case ID", "Tool Version")
    headersOK = True
    headerMsg = ""
    For hdrIdx = LBound(expectedHeaders) To UBound(expectedHeaders)
        If masterWs.Cells(1, hdrIdx + 1).Value <> expectedHeaders(hdrIdx) Then
            headersOK = False
            headerMsg = headerMsg & "- Col " & Split(masterWs.Cells(1, hdrIdx + 1).Address, "$")(1) & _
                " expected '" & expectedHeaders(hdrIdx) & "' but found '" & _
                masterWs.Cells(1, hdrIdx + 1).Value & "'" & vbCrLf
        End If
    Next hdrIdx

    If headersOK Then
        mRow = masterWs.Cells(masterWs.Rows.count, "A").End(xlUp).row + 1
        masterWs.Cells(mRow, 1).Value = Now
        masterWs.Cells(mRow, 2).Value = Environ("USERNAME")
        masterWs.Cells(mRow, 3).Value = ecmID
        masterWs.Cells(mRow, 4).Value = "2.4.1"
        If openedHidden Then
            masterWb.Close SaveChanges:=True
        Else
            masterWb.Save
        End If
    Else
        MsgBox "DIAGNOSTIC WARNING! HEADER MISMATCH ON " & targetSheet & vbCrLf & vbCrLf & _
               "The master tracker headers on " & targetSheet & " have been altered:" & vbCrLf & vbCrLf & _
               headerMsg & vbCrLf & _
               "Tracker data was NOT saved. Please notify the team lead.", vbCritical, "Diagnostic Failed"
        If openedHidden Then masterWb.Close SaveChanges:=False
    End If

CleanExit:
    On Error Resume Next
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    On Error GoTo 0
End Sub
