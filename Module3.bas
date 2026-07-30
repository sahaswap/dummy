Attribute VB_Name = "Module3"
Sub ExportToWord()
Dim wdApp As Object
Dim wdDoc As Object
Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim tagStr As String
Dim valStr As String
Dim masterShell As String
Dim shellRow As Long
Dim rng As Object

' --- NEW VARIABLES FOR SAVING ---
Dim ecmID As String
Dim AlertID As String
Dim desktopPath As String
Dim folderPath As String
Dim wordFileName As String
Dim CustName As String

' --- NEW VARIABLES FOR RFI TABLE & HEADER/FOOTER ---
Dim isRFI As Boolean
Dim sec As Object
Dim hdrRange As Object
Dim ftrRange As Object
Dim docText As String
Dim infoPos As Long
Dim promptPos As Long
Dim questionsText As String
Dim qlist() As String
Dim qArray() As String
Dim qCount As Long
Dim k As Long
Dim lineText As String
Dim matchNum As Boolean
Dim dotPos As Long
Dim numPart As String
Dim tableRng As Object
Dim wdTable As Object
Dim rIdx As Long
Dim c1 As Object
Dim c2 As Object

' 1. Point the macro to Sheet7 (where your mapping table lives)
Set ws = ThisWorkbook.Sheets("Sheet7")

' --- CAPTURE ID's FROM SHEET1 FOR NAMING ---
ecmID = ThisWorkbook.Sheets("Sheet1").Range("J10").Value
AlertID = ThisWorkbook.Sheets("Sheet1").Range("J11").Value

lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

' 2. Hunt for [MASTER_SHELL] in Column J
shellRow = 0
For i = 1 To lastRow
    If Trim(ws.Cells(i, "J").Value) = "[MASTER_SHELL]" Then
        shellRow = i
        Exit For
    End If
Next i

If shellRow = 0 Then
    MsgBox "Could not find [MASTER_SHELL] in Column J!", vbCritical
    Exit Sub
End If

' 3. Grab the template text from Column K of that row
masterShell = ws.Cells(shellRow, "K").Value

If masterShell = "" Or masterShell = "Select Template" Then
    MsgBox "Please select a template type on Sheet1 first!", vbExclamation
    Exit Sub
End If

' 4. Open Word
On Error Resume Next
Set wdApp = GetObject(, "Word.Application")
If wdApp Is Nothing Then
    Set wdApp = CreateObject("Word.Application")
End If
On Error GoTo 0

wdApp.Visible = True
Set wdDoc = wdApp.Documents.Add

' 5. Paste the template into Word
wdDoc.content.text = masterShell

' 6. Loop through Column J (Tags) and replace with Column K (Data)
For i = 1 To lastRow
    tagStr = ws.Cells(i, "J").Value
    
    If tagStr = "[Program Description]" Then
        valStr = ws.Cells(i, "K").Value
    Else
        valStr = ws.Cells(i, "K").text
    End If

    If tagStr <> "" And tagStr <> "[MASTER_SHELL]" And tagStr <> "Template Tag" Then
        Set rng = wdDoc.content
        
        With rng.Find
            .ClearFormatting
            .text = tagStr
            .Forward = True
            .Wrap = 0
            
            Do While .Execute = True
                rng.text = valStr
                rng.Collapse Direction:=0
            Loop
        End With
    End If
Next i

' Apply formatting to all narrative text: Justify alignment, 1.15 line spacing
With wdDoc.content.ParagraphFormat
    .Alignment = 3 ' wdAlignParagraphJustify
    .LineSpacingRule = 5 ' wdLineSpaceMultiple
    .LineSpacing = wdApp.LinesToPoints(1.15)
    .SpaceBefore = 0
    .SpaceAfter = 0
End With

' --- NEW: RFI FORMATTING DELEGATED TO LOCAL SUBROUTINE ---
isRFI = (UCase(Trim(ThisWorkbook.Sheets("Sheet1").Range("J6").Value)) = "RFI")
If isRFI Then
    FormatRFIDocument wdApp, wdDoc
End If

' --------------------------------------------------------
' 7. AUTO-CREATE FOLDER & SAVE WORD (.docx)
' --------------------------------------------------------
' Using your original shell path logic with backslash fixes
desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
folderPath = desktopPath & ecmID & "\"

' Create folder if it doesn't exist
If Len(Dir(folderPath, vbDirectory)) = 0 Then
    MkDir folderPath
End If

' Define file name
CustName = ThisWorkbook.Sheets("Sheet1").Range("J14").Value

If isRFI Then
    wordFileName = ecmID & "_" & AlertID & "_" & CustName & "_RFI QUESTIONS.docx"
Else
    wordFileName = ecmID & "_" & AlertID & "_" & CustName & "_Alert Write-Up.docx"
End If

' Save the file
On Error Resume Next
wdDoc.SaveAs2 fileName:=folderPath & wordFileName, FileFormat:=12
If Err.Number <> 0 Then
    MsgBox "Warning: Word Auto-save failed. Check if file is already open.", vbExclamation
    Err.Clear
End If
On Error GoTo 0

' Capture the full generated document text now, straight off the
' live Word doc, before anything downstream could touch it.
Dim generatedDocText As String
On Error Resume Next
generatedDocText = wdDoc.content.text
On Error GoTo 0

' Centralized audit ledger row - Register tab of this case's own
' Desktop\{ecmID}\{ecmID}_Audit_Log.xlsx. Detail here is just a
' pointer - the full text goes into that workbook's own Narrative
' tab via ArchiveNarrativeText below, not duplicated into this cell.
modAuditLog.LogAuditEvent ecmID:=ecmID, alertID:=AlertID, _
customerName:=CustName, _
counterparties:=modAuditLog.GetCounterpartyList(ThisWorkbook.Sheets("Sheet1")), _
eventType:=IIf(isRFI, "RFI Questions Generated", "Alert Write-Up Generated"), _
outputFile:=folderPath & wordFileName, _
toolVersion:="2.4.1", _
detail:="Full text archived in 'Narrative' tab of " & ecmID & "_Audit_Log.xlsx", _
notes:=""

' Writes the full text into that Narrative tab (latest version only
' - the Register row above already records every time this ran).
modAuditLog.ArchiveNarrativeText ecmID:=ecmID, fullText:=generatedDocText, _
docLabel:=wordFileName

' ==========================================
' 8. LIVE PUSH TO SHARED ONEDRIVE MASTER TRACKER (SHEET3) - HEADLESS GHOST MODE
' ==========================================
Dim masterPath As String, masterWb As Workbook, masterWs As Worksheet, pushWb As Workbook
Dim mRow As Long, wasAlreadyOpen As Boolean
Dim expectedHeaders As Variant, hdrIdx As Integer, headersOK As Boolean, headerMsg As String
Dim ghostApp As Object

masterPath = Environ("USERPROFILE") & "\OneDrive - Community Federal Savings Bank\Mohini Srivastava's files - L1 Beta\Beta 2.4_Feedbacks & Issues Encountered.xlsx"

' Check if the file is already open in the visible Excel window
wasAlreadyOpen = False
For Each pushWb In Application.Workbooks
If pushWb.Name = "Beta 2.4_Feedbacks & Issues Encountered.xlsx" Then
Set masterWb = pushWb
wasAlreadyOpen = True
Exit For
End If
Next pushWb

' Open silently in the background
If masterWb Is Nothing Then
Set ghostApp = CreateObject("Excel.Application")
ghostApp.Visible = False
ghostApp.DisplayAlerts = False
ghostApp.EnableEvents = False

On Error Resume Next
Set masterWb = ghostApp.Workbooks.Open(fileName:=masterPath, UpdateLinks:=False)
On Error GoTo 0
End If

' Process Data on SHEET3
If Not masterWb Is Nothing Then
If Not masterWb.ReadOnly Then
On Error Resume Next
Set masterWs = masterWb.Sheets("Sheet3") ' <--- TARGETING SHEET 3
On Error GoTo 0

    If Not masterWs Is Nothing Then
        expectedHeaders = Array("Date & Time", "Analyst ID", "ECM Case ID", "Tool Version")
        headersOK = True
        headerMsg = ""
        
        For hdrIdx = LBound(expectedHeaders) To UBound(expectedHeaders)
            If masterWs.Cells(1, hdrIdx + 1).Value <> expectedHeaders(hdrIdx) Then
                headersOK = False
                headerMsg = headerMsg & "- Col " & Split(masterWs.Cells(1, hdrIdx + 1).Address, "$")(1) & " expected '" & expectedHeaders(hdrIdx) & "' but found '" & masterWs.Cells(1, hdrIdx + 1).Value & "'" & vbCrLf
            End If
        Next hdrIdx
        
        If headersOK Then
            mRow = masterWs.Cells(masterWs.Rows.Count, "A").End(xlUp).Row + 1
            
            masterWs.Cells(mRow, 1).Value = Now
            masterWs.Cells(mRow, 2).Value = Environ("USERNAME")
            masterWs.Cells(mRow, 3).Value = ecmID
            masterWs.Cells(mRow, 4).Value = "2.4.1"
            
            If wasAlreadyOpen Then
                masterWb.Save
            Else
                masterWb.Close SaveChanges:=True
            End If
            
        Else
            MsgBox "DIAGNOSTIC WARNING! HEADER MISMATCH ON SHEET3" & vbCrLf & vbCrLf & _
                   "The master tracker headers on Sheet3 have been altered:" & vbCrLf & vbCrLf & _
                   headerMsg & vbCrLf & _
                   "Tracker data was NOT saved. Please notify the team lead.", vbCritical, "Diagnostic Failed"
            
            If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
        End If
    End If
Else
    ' File is locked by another user
    If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
End If
End If

' Destroy the ghost app
If Not ghostApp Is Nothing Then
ghostApp.Quit
Set ghostApp = Nothing
End If
' ==========================================

MsgBox "Export Complete! File saved to: " & folderPath, vbInformation

Set wdDoc = Nothing
Set wdApp = Nothing
End Sub

Public Sub FormatRFIDocument(ByVal wdApp As Object, ByVal wdDoc As Object)
    Dim sec As Object
    Dim hdrRange As Object
    Dim ftrRange As Object
    Dim docText As String
    Dim infoPos As Long
    Dim promptPos As Long
    Dim questionsText As String
    Dim qlist() As String
    Dim qArray() As String
    Dim qCount As Long
    Dim k As Long
    Dim lineText As String
    Dim matchNum As Boolean
    Dim dotPos As Long
    Dim numPart As String
    Dim tableRng As Object
    Dim wdTable As Object
    Dim rIdx As Long
    Dim qRow As Long
    Dim aRow As Long
    Dim c1 As Object
    Dim c2 As Object
    Dim bodyRng As Object
    Dim endRng As Object
    Dim titleRng As Object
    Dim p As Object
    Dim pIdx As Long
    Dim pText As String
    Dim r As Object
    Dim firstChar As String
    
    ' 0. Set Page Margins to 1.0 Inch (72 points)
    With wdDoc.PageSetup
        .TopMargin = 72
        .BottomMargin = 72
        .LeftMargin = 72
        .RightMargin = 72
    End With
    
    ' Set Corporate Styling for all body text: Arial 11 (Line spacing and alignment set globally)
    With wdDoc.content.ParagraphFormat
        .LineSpacingRule = 5 ' wdLineSpaceMultiple
        .LineSpacing = wdApp.LinesToPoints(1.15)
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
    wdDoc.content.Font.Name = "Arial"
    wdDoc.content.Font.Size = 11
    wdDoc.content.Font.Color = RGB(0, 0, 0)
    
    ' 1. Set page header and footer
    For Each sec In wdDoc.Sections
        ' Right-aligned header
        Set hdrRange = sec.Headers(1).Range
        hdrRange.text = "Financial Crimes Compliance Unit" & vbCrLf & "AML / TM Department"
        hdrRange.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
        hdrRange.ParagraphFormat.LineSpacingRule = 0
        hdrRange.ParagraphFormat.SpaceAfter = 0
        hdrRange.Font.Name = "Arial"
        hdrRange.Font.Size = 9
        hdrRange.Font.Color = RGB(0, 0, 0)
        
        ' Centered page numbers in footer
        Set ftrRange = sec.Footers(1).Range
        ftrRange.text = ""
        ftrRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
        ftrRange.ParagraphFormat.LineSpacingRule = 0
        ftrRange.ParagraphFormat.SpaceAfter = 0
        ftrRange.Font.Name = "Arial"
        ftrRange.Font.Size = 10
        ftrRange.Font.Color = RGB(0, 0, 0)
        
        ' Add page fields dynamically (33 = wdFieldPage, 26 = wdFieldNumPages)
        ftrRange.Fields.Add Range:=ftrRange, Type:=33
        
        Set ftrRange = sec.Footers(1).Range
        ftrRange.Collapse Direction:=0
        ftrRange.text = " of "
        
        Set ftrRange = sec.Footers(1).Range
        ftrRange.Collapse Direction:=0
        ftrRange.Fields.Add Range:=ftrRange, Type:=26
        
        Set ftrRange = sec.Footers(1).Range
        ftrRange.InsertBefore "Page "
    Next sec
    
    ' 2. Strip compliance header from document body (including surrounding whitespace/tabs)
    Set bodyRng = wdDoc.content
    With bodyRng.Find
        .ClearFormatting
        .text = "Financial Crimes Compliance Unit"
        .Replacement.text = ""
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    Set bodyRng = wdDoc.content
    With bodyRng.Find
        .ClearFormatting
        .text = "AML / TM Department"
        .Replacement.text = ""
        .Execute Replace:=2
    End With
    
    ' Clean up starting whitespace in body (tabs, spaces, carriage returns)
    Do While Len(wdDoc.content.text) > 0 And (Left(wdDoc.content.text, 1) = vbCr Or Left(wdDoc.content.text, 1) = vbLf Or Left(wdDoc.content.text, 1) = vbTab Or Left(wdDoc.content.text, 1) = " ")
        wdDoc.Range(0, 1).Delete
    Loop
    
    ' 3. Insert and Format Title centered and underlined at top
    If InStr(wdDoc.content.text, "Request for Information (RFI)") <> 1 Then
        Set titleRng = wdDoc.Range(0, 0)
        titleRng.text = "Request for Information (RFI)" & vbCrLf
    Else
        Set titleRng = wdDoc.Paragraphs(1).Range
    End If
    titleRng.ParagraphFormat.Alignment = 1 ' Center
    titleRng.ParagraphFormat.SpaceBefore = 0
    titleRng.ParagraphFormat.SpaceAfter = 12
    titleRng.ParagraphFormat.LineSpacingRule = 0 ' Single spacing
    titleRng.Font.Name = "Arial"
    titleRng.Font.Size = 11
    titleRng.Font.bold = True
    titleRng.Font.Underline = 1 ' Underline
    titleRng.Font.Color = RGB(0, 0, 0)
    
    ' 4. Clean up mail routing keywords
    Set bodyRng = wdDoc.content
    With bodyRng.Find
        .ClearFormatting
        .text = "send cfsbencrypt"
        .Replacement.text = ""
        .Execute Replace:=2
    End With
    
    ' 5. Convert text bullets to native Word bullet points
    For pIdx = wdDoc.Paragraphs.Count To 1 Step -1
        Set p = wdDoc.Paragraphs(pIdx)
        pText = Trim(p.Range.text)
        If Len(pText) > 0 Then
            firstChar = Left(pText, 1)
            If firstChar = ChrW(8226) Or firstChar = Chr(149) Then
                Set r = p.Range
                Do While r.Characters.Count > 0
                    firstChar = r.Characters(1).text
                    If firstChar = ChrW(8226) Or firstChar = Chr(149) Or firstChar = " " Or firstChar = vbTab Then
                        r.Characters(1).Delete
                    Else
                        Exit Do
                    End If
                Loop
                p.Range.ListFormat.ApplyBulletDefault
                ' Apply styling: Arial 11, line spacing 1
                p.Range.Font.Name = "Arial"
                p.Range.Font.Size = 11
                p.Range.ParagraphFormat.LineSpacingRule = 0
                p.Range.ParagraphFormat.SpaceAfter = 0
            End If
        End If
    Next pIdx
    
    ' 6. Extract questions list and build the Word Table
    docText = wdDoc.content.text
    
    infoPos = InStr(docText, "Information Requested:")
    promptPos = InStr(docText, "Your prompt response")
    
    If infoPos > 0 And promptPos > infoPos Then
        questionsText = Mid(docText, infoPos, promptPos - infoPos)
        
        ' Normalize line breaks to vbCr before splitting
        questionsText = Replace(questionsText, vbCrLf, vbCr)
        questionsText = Replace(questionsText, vbLf, vbCr)
        questionsText = Replace(questionsText, Chr(11), vbCr)
        
        ' Parse questions
        qlist = Split(questionsText, vbCr)
        
        ReDim qArray(1 To 1)
        qCount = 0
        
        For k = LBound(qlist) To UBound(qlist)
            lineText = Trim(qlist(k))
            lineText = Replace(lineText, vbLf, "")
            lineText = Replace(lineText, vbTab, "")
            
            matchNum = False
            If Len(lineText) > 2 Then
                dotPos = InStr(lineText, ".")
                If dotPos > 1 And dotPos < 5 Then
                    numPart = Left(lineText, dotPos - 1)
                    If IsNumeric(numPart) Then
                        matchNum = True
                        lineText = Trim(Mid(lineText, dotPos + 1))
                    End If
                End If
            End If
            
            If matchNum And lineText <> "" Then
                qCount = qCount + 1
                ReDim Preserve qArray(1 To qCount)
                
                ' Replace slash with newline ONLY for Entity / Individual
                If InStr(lineText, "Entity:") > 0 And InStr(lineText, "Individual:") > 0 Then
                    lineText = Replace(lineText, " / Individual:", vbCrLf & "Individual:")
                    lineText = Replace(lineText, "/ Individual:", vbCrLf & "Individual:")
                    lineText = Replace(lineText, " /Individual:", vbCrLf & "Individual:")
                    lineText = Replace(lineText, "/Individual:", vbCrLf & "Individual:")
                End If
                
                qArray(qCount) = lineText
            End If
        Next k
        
        If qCount > 0 Then
            ' Inject table replacing the raw text list
            Set tableRng = wdDoc.content
            
            With tableRng.Find
                .ClearFormatting
                .text = "Information Requested:"
                If .Execute Then
                    Set endRng = wdDoc.content
                    With endRng.Find
                        .ClearFormatting
                        .text = "Your prompt response"
                        If .Execute Then
                            tableRng.End = endRng.Start
                            tableRng.text = ""
                            
                            ' Create table with double rows (1 Question + 1 Answer row per question)
                            Set wdTable = wdDoc.Tables.Add(Range:=tableRng, NumRows:=qCount * 2, NumColumns:=2)
                            
                            ' Table Border Formatting (full grid in gray)
                            wdTable.Borders.Enable = True
                            wdTable.Borders(-1).Color = RGB(127, 127, 127) ' Top
                            wdTable.Borders(-2).Color = RGB(127, 127, 127) ' Left
                            wdTable.Borders(-3).Color = RGB(127, 127, 127) ' Bottom
                            wdTable.Borders(-4).Color = RGB(127, 127, 127) ' Right
                            wdTable.Borders(-5).Color = RGB(127, 127, 127) ' Inside Horizontal
                            wdTable.Borders(-6).Color = RGB(127, 127, 127) ' Inside Vertical
                            wdTable.Borders(-6).LineStyle = 1              ' Visible vertical lines
                            
                            ' Columns widths: (Col 1 = 0.5 inches / 36 points, Col 2 = 6.0 inches / 432 points)
                            wdTable.Columns(1).Width = 36
                            wdTable.Columns(2).Width = 432
                            
                            ' Padding
                            wdTable.TopPadding = 5
                            wdTable.BottomPadding = 5
                            wdTable.LeftPadding = 8
                            wdTable.RightPadding = 8
                            
                            ' Populate Cells
                            For rIdx = 1 To qCount
                                qRow = 2 * rIdx - 1
                                aRow = 2 * rIdx
                                
                                ' -- Question Row --
                                Set c1 = wdTable.cell(qRow, 1)
                                Set c2 = wdTable.cell(qRow, 2)
                                
                                c1.Range.text = "(" & rIdx & ")"
                                c1.Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                                c1.Range.ParagraphFormat.SpaceAfter = 0
                                c1.Range.ParagraphFormat.LineSpacingRule = 0
                                c1.Range.Font.bold = True
                                c1.Range.Font.Name = "Arial"
                                c1.Range.Font.Size = 11
                                c1.Range.Font.Color = RGB(0, 0, 0)
                                c1.Shading.BackgroundPatternColor = RGB(218, 233, 247) ' #DAE9F7 exact blue
                                
                                c2.Range.text = qArray(rIdx)
                                c2.Range.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                                c2.Range.ParagraphFormat.SpaceAfter = 0
                                c2.Range.ParagraphFormat.LineSpacingRule = 0
                                c2.Range.Font.bold = False
                                c2.Range.Font.Name = "Arial"
                                c2.Range.Font.Size = 11
                                c2.Range.Font.Color = RGB(0, 0, 0)
                                c2.Shading.BackgroundPatternColor = RGB(218, 233, 247) ' #DAE9F7 exact blue
                                
                                ' -- Answer Row --
                                Set c1 = wdTable.cell(aRow, 1)
                                Set c2 = wdTable.cell(aRow, 2)
                                
                                c1.Range.text = ""
                                c1.Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                                c1.Range.ParagraphFormat.SpaceAfter = 0
                                c1.Range.ParagraphFormat.LineSpacingRule = 0
                                c1.Range.Font.Name = "Arial"
                                c1.Range.Font.Size = 11
                                c1.Shading.BackgroundPatternColor = RGB(255, 255, 255) ' White
                                
                                c2.Range.text = ""
                                c2.Range.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                                c2.Range.ParagraphFormat.SpaceAfter = 0
                                c2.Range.ParagraphFormat.LineSpacingRule = 0
                                c2.Range.Font.Name = "Arial"
                                c2.Range.Font.Size = 11
                                c2.Range.Font.Color = RGB(0, 0, 0)
                                c2.Shading.BackgroundPatternColor = RGB(255, 255, 255) ' White
                            Next rIdx
                        End If
                    End With
                End If
            End With
        End If
    End If
End Sub
