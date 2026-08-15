Attribute VB_Name = "modRFIGenerate"
'==================================================================
' modRFIGenerate  -  RFI data layer + (later) orchestrator
'
' This file currently provides the DATA foundations for RFI:
'   - RFIData type (the dynamic fields a generated RFI needs)
'   - GetSheet7Tag    : reuse existing Sheet7 [tag] -> value map
'   - GatherRFIData   : build an RFIData record from Sheet1 + Sheet7
'   - SanitizeFileNamePart : make a string safe for a Windows filename
'
' The orchestrator (GenerateRFI), SaveRFIDocument, and PushRFITrackerRow
' are added in later tasks (4-5).
'
' Tag mapping (finalized): [Counterparties] fills BOTH the credit
' (ORIGINATOR) and debit (COUNTERPARTY) example slots. Due Date,
' Coordinator Name and the Common "Program Contact" line are NOT
' gathered - they stay as manual placeholders in the document.
'==================================================================
Option Explicit

' --- Dynamic fields gathered per case ---
Public Type RFIData
    customerName As String     ' Sheet1!J13
    AccountNumber As String    ' Sheet7 [Account Numbers]
    DateStart As String        ' Sheet7 [Start Date]
    DateEnd As String          ' Sheet7 [End Date]
    NumCr As String            ' Sheet7 [Num Cr]   - credit example line
    CrTotal As String          ' Sheet7 [Cr]       - credit example line
    NumDr As String            ' Sheet7 [Num Dr]   - debit example line
    DrTotal As String          ' Sheet7 [Dr]       - debit example line
    counterparties As String   ' Sheet7 [Counterparties] - both ORIGINATOR & COUNTERPARTY
    CaseNumber As String       ' Sheet1!J9
    AlertNumber As String      ' Sheet1!J10
End Type

' Version stamp written to the tracker row. Module-level declarations
' must sit ABOVE the first procedure, otherwise VBA raises "only
' comments may appear after End Sub/Function/Property".
Private Const RFI_TOOL_VERSION As String = "3.5"

'------------------------------------------------------------------
' GetSheet7Tag
'   PRE:  tagName is a bracketed tag exactly as stored in Sheet7 col J,
'         e.g. "[Total Amount]".
'   POST: returns the displayed .Text of the matching col-K cell, or ""
'         if the tag is not found or the cell holds an error. Never raises.
'------------------------------------------------------------------
Public Function GetSheet7Tag(ByVal tagName As String) As String
    Dim ws As Worksheet, lastRow As Long, i As Long

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Sheets("Sheet7")
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).row

    For i = 1 To lastRow
        If Trim$(CStr(ws.Cells(i, "J").Value)) = tagName Then
            If Not isError(ws.Cells(i, "K").Value) Then
                GetSheet7Tag = ws.Cells(i, "K").text
            End If
            Exit Function
        End If
    Next i
    Exit Function

CleanFail:
    GetSheet7Tag = ""
End Function

'------------------------------------------------------------------
' GatherRFIData
'   PRE:  Sheet1 input panel populated; Sheet7 aggregate formulas resolved.
'   POST: returns an RFIData record. Missing/errored aggregates default to "".
'         Due Date / Coordinator / Program Contact are intentionally NOT set
'         (they remain manual placeholders in the generated document).
'------------------------------------------------------------------
Public Function GatherRFIData() As RFIData
    Dim d As RFIData, ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    d.customerName = CStr(ws.Range("J13").Value)
    d.CaseNumber = CStr(ws.Range("J9").Value)
    d.AlertNumber = CStr(ws.Range("J10").Value)

    d.AccountNumber = GetSheet7Tag("[Account Numbers]")
    d.DateStart = GetSheet7Tag("[Start Date]")
    d.DateEnd = GetSheet7Tag("[End Date]")
    d.NumCr = GetSheet7Tag("[Num Cr]")
    d.CrTotal = GetSheet7Tag("[Cr]")
    d.NumDr = GetSheet7Tag("[Num Dr]")
    d.DrTotal = GetSheet7Tag("[Dr]")
    d.counterparties = GetSheet7Tag("[Counterparties]")

    GatherRFIData = d
End Function

'------------------------------------------------------------------
' SanitizeFileNamePart
'   PRE:  any string (may contain illegal filename chars / line breaks).
'   POST: returns a trimmed string safe to use as part of a Windows file
'         name: illegal chars and CR/LF/TAB replaced with spaces, runs of
'         spaces collapsed. Returns "" for empty/space-only input.
'------------------------------------------------------------------
Public Function SanitizeFileNamePart(ByVal text As String) As String
    Dim bad As Variant, i As Long

    text = Replace(text, Chr(10), " ")   ' LF
    text = Replace(text, Chr(13), " ")   ' CR
    text = Replace(text, Chr(9), " ")    ' TAB

    bad = Array("""", "/", "\", ":", "*", "?", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        text = Replace(text, bad(i), " ")
    Next i

    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop

    SanitizeFileNamePart = Trim$(text)
End Function

'==================================================================
' ORCHESTRATOR + SAVE + TRACKER  (tasks 4 & 5)
'==================================================================

' Tracker path/name come from modConfig (single source of truth), so
' this RFI push lands in the same shared workbook as ExportToWord.

'------------------------------------------------------------------
' GenerateRFI  -  entry point for an RFI case (called by the router).
'   PRE:  Sheet1!J5 = "RFI".
'   POST: builds the RFI Word doc, saves it to the case folder, pushes a
'         tracker row. Aborts with a MsgBox (no file) on any validation failure.
'------------------------------------------------------------------
Public Sub GenerateRFI()
    Dim ws As Worksheet, tpl As RFITemplate, data As RFIData
    Dim wdDoc As Object

    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' 1. Validate required inputs
    If Trim$(CStr(ws.Range("J9").Value)) = "" Then
        MsgBox "Missing ECM Case ID (J9). Please fill it before generating an RFI.", _
               vbExclamation, "RFI": Exit Sub
    End If
    If Trim$(CStr(ws.Range("J10").Value)) = "" Then
        MsgBox "Missing Alert ID (J10). Please fill it before generating an RFI.", _
               vbExclamation, "RFI": Exit Sub
    End If
    If Trim$(CStr(ws.Range("J13").Value)) = "" Then
        MsgBox "Missing Customer Name (J13). Please fill it before generating an RFI.", _
               vbExclamation, "RFI": Exit Sub
    End If

    ' 2. Resolve the template from J6
    tpl = GetRFITemplate(CStr(ws.Range("J6").Value))
    If tpl.Name = "" Then
        MsgBox "For an RFI, the Template (J6) must be Wise, Airwallex, or Common.", _
               vbExclamation, "RFI": Exit Sub
    End If

    ' 3. Gather dynamic fields (reusing Sheet7 aggregates)
    data = GatherRFIData()

    ' 4. Build the Word document
    Set wdDoc = BuildRFIDocument(tpl, data)
    If wdDoc Is Nothing Then Exit Sub        ' builder already messaged the user

    ' 5. Save + tracker
    Dim savedPath As String
    savedPath = SaveRFIDocument(wdDoc, CStr(ws.Range("J9").Value), CStr(ws.Range("J13").Value))
    PushRFITrackerRow CStr(ws.Range("J9").Value)

    ' Centralized audit ledger row - Register tab of this case's own
    ' Desktop\{ECMID}\{ECMID}_Audit_Log.xlsx. Skipped entirely if the
    ' save itself failed - nothing to point at yet. Detail here is
    ' just a pointer - the full text goes into that workbook's own
    ' Narrative tab via ArchiveNarrativeText below.
    If savedPath <> "" Then
        Dim generatedRFIText As String
        On Error Resume Next
        generatedRFIText = wdDoc.content.text
        On Error GoTo 0

        Dim ecmIDForLog As String
        ecmIDForLog = CStr(ws.Range("J9").Value)

        modAuditLog.LogAuditEvent ecmID:=ecmIDForLog, _
            AlertID:=CStr(ws.Range("J10").Value), _
            customerName:=CStr(ws.Range("J13").Value), _
            counterparties:=modAuditLog.GetCounterpartyList(ws), _
            eventType:="RFI Generated", _
            outputFile:=savedPath, _
            toolVersion:=RFI_TOOL_VERSION, _
            notes:="template=" & tpl.Name, _
            detail:="Full text archived in 'Narrative' tab of " & ecmIDForLog & "_Audit_Log.xlsx"

        modAuditLog.ArchiveNarrativeText ecmID:=ecmIDForLog, fullText:=generatedRFIText, _
            docLabel:=Mid$(savedPath, InStrRev(savedPath, "\") + 1)
    End If

    MsgBox "RFI (" & tpl.Name & ") generated and saved to the case folder.", _
           vbInformation, "RFI Ready"
End Sub

'------------------------------------------------------------------
' SaveRFIDocument
'   Saves to Desktop\{ECMID}\{ECMID}_{CustomerName}_RFI Question.docx (FileFormat 12).
'   Folder created if missing; save errors warned but non-fatal.
'   POST: returns the full path saved to, or "" if the save failed
'         (so the caller can skip/annotate the audit log entry).
'------------------------------------------------------------------
Private Function SaveRFIDocument(ByRef wdDoc As Object, ByVal ecmID As String, ByVal customerName As String) As String
    Dim desktopPath As String, folderPath As String, fileName As String
    Dim ecmClean As String, custClean As String

    ecmClean = SanitizeFileNamePart(ecmID)
    custClean = SanitizeFileNamePart(customerName)

    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    folderPath = desktopPath & ecmClean & "\"
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath

    fileName = ecmClean & "_" & custClean & "_RFI Question.docx"

    On Error Resume Next
    wdDoc.SaveAs2 fileName:=folderPath & fileName, FileFormat:=12
    If Err.Number <> 0 Then
        MsgBox "Warning: RFI auto-save failed (the file may already be open).", _
               vbExclamation, "RFI"
        Err.Clear
        SaveRFIDocument = ""
    Else
        SaveRFIDocument = folderPath & fileName
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------
' PushRFITrackerRow
'   Appends a run row to the shared tracker (Sheet3) using the headless-ghost
'   pattern from ExportToWord. All failures are non-fatal (the doc is saved).
'------------------------------------------------------------------
Private Sub PushRFITrackerRow(ByVal ecmID As String)
    Dim masterPath As String, masterWb As Workbook, masterWs As Worksheet, pushWb As Workbook
    Dim mRow As Long, wasAlreadyOpen As Boolean
    Dim expectedHeaders As Variant, hdrIdx As Integer, headersOK As Boolean
    Dim ghostApp As Object

    masterPath = TrackerFile()

    wasAlreadyOpen = False
    For Each pushWb In Application.Workbooks
        If pushWb.Name = TrackerFileName() Then
            Set masterWb = pushWb
            wasAlreadyOpen = True
            Exit For
        End If
    Next pushWb

    If masterWb Is Nothing Then
        Set ghostApp = CreateObject("Excel.Application")
        ghostApp.Visible = False
        ghostApp.DisplayAlerts = False
        ghostApp.EnableEvents = False
        On Error Resume Next
        Set masterWb = ghostApp.Workbooks.Open(fileName:=masterPath, UpdateLinks:=False)
        On Error GoTo 0
    End If

    If Not masterWb Is Nothing Then
        If Not masterWb.ReadOnly Then
            On Error Resume Next
            Set masterWs = masterWb.Sheets("Sheet3")
            On Error GoTo 0

            If Not masterWs Is Nothing Then
                expectedHeaders = Array("Date & Time", "Analyst ID", "ECM Case ID", "Tool Version")
                headersOK = True
                For hdrIdx = LBound(expectedHeaders) To UBound(expectedHeaders)
                    If masterWs.Cells(1, hdrIdx + 1).Value <> expectedHeaders(hdrIdx) Then
                        headersOK = False
                        Exit For
                    End If
                Next hdrIdx

                If headersOK Then
                    mRow = masterWs.Cells(masterWs.Rows.count, "A").End(xlUp).row + 1
                    masterWs.Cells(mRow, 1).Value = Now
                    masterWs.Cells(mRow, 2).Value = Environ("USERNAME")
                    masterWs.Cells(mRow, 3).Value = ecmID
                    masterWs.Cells(mRow, 4).Value = RFI_TOOL_VERSION
                    If wasAlreadyOpen Then
                        masterWb.Save
                    Else
                        masterWb.Close SaveChanges:=True
                    End If
                Else
                    If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
                End If
            End If
        Else
            If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
        End If
    End If

    If Not ghostApp Is Nothing Then
        ghostApp.Quit
        Set ghostApp = Nothing
    End If
End Sub

