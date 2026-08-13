Attribute VB_Name = "modRFIWord"
'==================================================================
' modRFIWord  -  In-macro Word document builder for RFI emails
'
' Builds the RFI Word document entirely via the Word object model
' (late-bound), so no external .docx template files are needed. The
' question table is constructed programmatically (Tables.Add) with
' question text baked in and the response column left blank.
'
' Public entry: BuildRFIDocument(tpl, data) As Object  (a Word.Document)
' Consumed by: modRFIGenerate.GenerateRFI
'==================================================================
Option Explicit

' --- Word enum values needed under late binding (no library reference) ---
Private Const wdAlignParagraphLeft   As Long = 0
Private Const wdAlignParagraphRight  As Long = 2
Private Const wdCollapseEnd          As Long = 0
Private Const wdLineStyleSingle      As Long = 1
Private Const wdPreferredWidthPercent As Long = 2
Private Const wdListBullet           As Long = 2
Private Const wdStory                As Long = 6

'------------------------------------------------------------------
' BuildRFIDocument
'   PRE:  tpl fully populated (Questions 1-based, non-empty); data gathered.
'   POST: returns a visible Word.Document laid out as the RFI email.
'         Returns Nothing (after a MsgBox) if Word automation fails.
'------------------------------------------------------------------
Public Function BuildRFIDocument(ByRef tpl As RFITemplate, ByRef data As RFIData) As Object
    Dim wdApp As Object, wdDoc As Object

    On Error GoTo WordFail
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo WordFail

    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add

    ' (a) page margins (1 inch all round)
    With wdDoc.PageSetup
        .TopMargin = wdApp.InchesToPoints(1)
        .BottomMargin = wdApp.InchesToPoints(1)
        .LeftMargin = wdApp.InchesToPoints(1)
        .RightMargin = wdApp.InchesToPoints(1)
    End With

    ' (b) right-aligned header block
    AppendPara wdDoc, RFI_HEADER_LINE1, wdAlignParagraphRight, True
    AppendPara wdDoc, RFI_HEADER_LINE2, wdAlignParagraphRight, False
    AppendPara wdDoc, "", wdAlignParagraphLeft, False

    ' (c) To / Cc / Subject block (per template)
    AppendPara wdDoc, "To: " & tpl.ToLine, wdAlignParagraphLeft, False
    AppendPara wdDoc, "Cc: " & tpl.CcLine, wdAlignParagraphLeft, False
    AppendPara wdDoc, "Subject: " & BuildSubject(tpl, data), wdAlignParagraphLeft, False
    AppendPara wdDoc, "", wdAlignParagraphLeft, False

    ' (d) greeting + intro
    AppendPara wdDoc, "Good Morning / Afternoon,", wdAlignParagraphLeft, False
    AppendPara wdDoc, "CFSB requires additional information regarding the transaction(s) " & _
        "and/or customer relationship(s) noted below:", wdAlignParagraphLeft, False

    ' (e) customer bullet + example credit / debit transaction lines
    AppendBullet wdDoc, "We are reviewing transactions for your customer, " & _
        data.customerName & " (account # " & data.AccountNumber & ")."
    AppendPara wdDoc, "Ex: Credit Transactions", wdAlignParagraphLeft, False
    AppendBullet wdDoc, "Between " & data.DateStart & " to " & data.DateEnd & _
        " there are " & data.NumCr & " incoming credit transactions totaling " & _
        data.CrTotal & " from " & data.counterparties & ". The purpose of this " & _
        "activity and relationship between the customer and the counterparty appears to be unknown."
    AppendPara wdDoc, "Ex: Debit Transactions", wdAlignParagraphLeft, False
    AppendBullet wdDoc, "Between " & data.DateStart & " to " & data.DateEnd & _
        " there are " & data.NumDr & " outgoing debit transactions totaling " & _
        data.DrTotal & " going to " & data.counterparties & ". The purpose of this " & _
        "activity and relationship between the customer and the counterparty appears to be unknown."
    AppendPara wdDoc, "", wdAlignParagraphLeft, False
    AppendPara wdDoc, "To proceed appropriately, we kindly request additional information " & _
        "regarding your customer and the transactions outlined below on or before close of " & _
        "business day, [RFI Due Date].", wdAlignParagraphLeft, False

    ' (f) numbered 2-column question table (response column blank)
    BuildQuestionTable wdDoc, tpl.questions

    ' (g) closing block
    AppendPara wdDoc, "", wdAlignParagraphLeft, False
    AppendPara wdDoc, "Your prompt response will allow us to complete our review in a " & _
        "timely manner.", wdAlignParagraphLeft, False
    AppendPara wdDoc, "", wdAlignParagraphLeft, False
    AppendPara wdDoc, "Thank you,", wdAlignParagraphLeft, False
    AppendPara wdDoc, "[Deloitte RFI Coordinator Name]", wdAlignParagraphLeft, False
    AppendPara wdDoc, RFI_ENCRYPT_TAG, wdAlignParagraphLeft, False

    Set BuildRFIDocument = wdDoc
    Exit Function

WordFail:
    MsgBox "Could not start Microsoft Word to build the RFI:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "RFI - Word Error"
    Set BuildRFIDocument = Nothing
End Function

'------------------------------------------------------------------
' BuildQuestionTable
'   PRE:  questions is 1-based, non-empty (6 for Wise/Common, 10 for Airwallex).
'   POST: an (N+1) x 2 bordered table is appended; bold header row,
'         one row per question prefixed "r. ", response column (col 2) blank.
'------------------------------------------------------------------
Private Sub BuildQuestionTable(ByRef wdDoc As Object, ByRef questions() As String)
    Dim rng As Object, tbl As Object
    Dim lo As Long, hi As Long, n As Long, r As Long

    lo = LBound(questions)
    hi = UBound(questions)
    n = hi - lo + 1
    If n < 1 Then Exit Sub

    Set rng = wdDoc.content
    rng.Collapse wdCollapseEnd

    Set tbl = wdDoc.Tables.Add(Range:=rng, NumRows:=n + 1, NumColumns:=2)

    With tbl
        .Borders.OutsideLineStyle = wdLineStyleSingle
        .Borders.InsideLineStyle = wdLineStyleSingle
        .PreferredWidthType = wdPreferredWidthPercent
        .Columns(1).PreferredWidth = 70
        .Columns(2).PreferredWidth = 30

        ' header row
        .cell(1, 1).Range.text = "Information Requested"
        .cell(1, 2).Range.text = "Response"
        .Rows(1).Range.Font.bold = True

        ' one row per question; response column intentionally blank
        For r = 1 To n
            .cell(r + 1, 1).Range.text = r & ". " & questions(lo + r - 1)
        Next r
    End With
End Sub

'------------------------------------------------------------------
' AppendPara - append a paragraph at end of doc with alignment/bold.
'------------------------------------------------------------------
Private Sub AppendPara(ByRef wdDoc As Object, ByVal text As String, _
                       ByVal align As Long, ByVal bold As Boolean)
    Dim rng As Object
    Set rng = wdDoc.content
    rng.Collapse wdCollapseEnd
    rng.Font.bold = bold
    rng.ParagraphFormat.Alignment = align
    rng.text = text & vbCr
End Sub

'------------------------------------------------------------------
' AppendBullet - append a bulleted line at end of doc.
'------------------------------------------------------------------
Private Sub AppendBullet(ByRef wdDoc As Object, ByVal text As String)
    Dim rng As Object
    Set rng = wdDoc.content
    rng.Collapse wdCollapseEnd
    rng.text = text & vbCr
    On Error Resume Next
    rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
    rng.ListFormat.ApplyBulletDefault
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
' BuildSubject - substitute {CUSTOMER}/{CASE}/{ALERT} in the template format.
'------------------------------------------------------------------
Private Function BuildSubject(ByRef tpl As RFITemplate, ByRef data As RFIData) As String
    Dim s As String
    s = tpl.SubjectFormat
    s = Replace(s, "{CUSTOMER}", data.customerName)
    s = Replace(s, "{CASE}", data.CaseNumber)
    s = Replace(s, "{ALERT}", data.AlertNumber)
    BuildSubject = s
End Function
