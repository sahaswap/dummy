Option Explicit

' ================================================================
' modSetCommonRFITemplate - ONE-TIME content update.
'
' Overwrites the Common RFI template text in Template_DB!B9 with the new
' approved layout: single combined example, 5 questions, "CFSB" sign-off,
' no coordinator name / encrypt tag.
'
' The compliance header stays at the top of the text on purpose:
' FormatRFIDocument strips it from the document BODY, and (with the
' Module3 change) no page header/footer is added for Common - so it
' disappears entirely, matching the approved format.
'
' Tags in [brackets] are replaced by ExportToWord from Sheet7. Literal
' placeholders ([Program Contact from list], [INSERT DESCRIPTION...],
' "ABC Corporation", "Month 22, 202X") are left for the analyst to fill.
'
' Run SetCommonRFITemplate once. Wise / Airwallex rows are untouched.
' ================================================================
Sub SetCommonRFITemplate()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Template_DB")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Template_DB sheet not found.", vbCritical
        Exit Sub
    End If

    Dim L As String, B As String, Q As String
    L = vbLf                 ' in-cell line break
    B = ChrW$(8226) & " "    ' bullet + space
    Q = Chr$(34)             ' double quote

    Dim t As String
    t = vbTab & vbTab & vbTab & "Financial Crimes Compliance Unit" & L
    t = t & vbTab & vbTab & vbTab & "AML / TM Department" & L & L
    t = t & "To: [Program Contact from list]" & L
    t = t & "Cc: AMLRequests@cfsb.com" & L
    t = t & "Subject: CFSB Request for Information - [Subject Name] " & _
            "(Velocity CASE #[ECM ID] / ALERT# [Alert ID])" & L & L
    t = t & "Good Morning / Afternoon" & L & L
    t = t & "CFSB requires additional information regarding the transaction(s) " & _
            "and/or customer relationship(s) noted below." & L & L
    t = t & B & "We are reviewing transactions for your customer, [Subject Name] " & _
            "(account # [Account Numbers]) [INSERT DESCRIPTION OF THE ACTIVITY YOU ARE " & _
            "REVIEWING NEED ADDITIONAL INFORMATION ABOUT]" & L & L
    t = t & "Ex: Between [Start Date] to [End Date] there are [Num Cr] credit transactions " & _
            "totaling to $[Cr] from [Counterparties] and there are [Num Dr] debit transactions " & _
            "totaling $[Dr] to Outgoing wire with reference to " & Q & "ABC Corporation" & Q & _
            ". The purpose of this activity with the counterparties appears unknown." & L & L
    t = t & "To proceed appropriately, we kindly request additional information regarding your " & _
            "customer and the transactions outlined below on or before close of business " & _
            "Month 22, 202X." & L & L
    t = t & "Information Requested:" & L
    t = t & "1. Ultimate Originators and Beneficiaries of the noted transactions" & L
    t = t & "2. Purpose of the Noted transactions" & L
    t = t & "3. Relationship between CUSTOMER and COUNTERPARTY" & L
    t = t & "4. Do you expect these transactions to continue?" & L
    t = t & "5. Please note any other relevant information / comments below." & L
    t = t & "Your prompt response will allow us to complete our review in a timely manner." & L & L
    t = t & "Thank you," & L
    t = t & "CFSB"

    Dim wasProt As Boolean
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect Password:="p7ss"
    On Error GoTo 0

    ws.Range("B9").Value = t

    On Error Resume Next
    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0

    MsgBox "Common RFI template written to Template_DB!B9." & vbCrLf & vbCrLf & _
           "Generate a Narrative with Decision = RFI and Template = Common to check it.", _
           vbInformation, "Common Template Updated"
End Sub
