Attribute VB_Name = "modSetNonEscTemplate"
Option Explicit

' ================================================================
' modSetNonEscTemplate - ONE-TIME content update.
'
' Replaces the Non-Escalation narrative (Template_DB row
' "Non-Escalation_Closure", which Sheet7's [MASTER_SHELL] uses for
' Non-Escalation with the Wise or Fluz template) with the combined
' Individual / Entity version. Where the two versions differ, both are
' kept, as "[Option 1 - Individual]" / "[Option 2 - Entity]" pairs or as
' "<Individual: ... / Entity: ...>" choices, and the analyst deletes the
' one that doesn't apply - the same convention as the Escalation
' templates' "[Option 1 - ...]" blocks.
'
' Tags in [brackets] are filled by ExportToWord from Sheet7. Text in
' <angle brackets>, and the bracketed section labels ([External
' Research], [Negative News], ...), are left for the analyst.
'
' Also adds a [Country Risk] tag to Sheet7 for the "Country Risk Rating"
' line: Sheet1's Excel Risk Class (T27) with the country (G27), e.g.
' "Medium (CYPRUS)".
'
' The text being replaced is kept on Template_DB as a new row,
' "Non-Escalation_Closure (previous)", so it can be put back. The Crypto
' Non-Escalation template is not touched.
'
' Run SetNonEscalationTemplate once.
' ================================================================
Sub SetNonEscalationTemplate()
    Dim wsT As Worksheet, ws7 As Worksheet
    Dim r As Long, keyRow As Long, lastRow As Long, tagRow As Long
    Dim oldText As String

    On Error Resume Next
    Set wsT = ThisWorkbook.Sheets("Template_DB")
    Set ws7 = ThisWorkbook.Sheets("Sheet7")
    On Error GoTo 0
    If wsT Is Nothing Or ws7 Is Nothing Then
        MsgBox "Template_DB or Sheet7 not found.", vbCritical
        Exit Sub
    End If

    ' --- the template row, found by its key ---
    lastRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
    For r = 1 To lastRow
        If Trim$(CStr(wsT.Cells(r, "A").Value)) = "Non-Escalation_Closure" Then keyRow = r: Exit For
    Next r
    If keyRow = 0 Then
        MsgBox "Template_DB has no ""Non-Escalation_Closure"" row, so nothing was changed.", vbCritical
        Exit Sub
    End If

    Dim L As String, A As String, LQ As String, RQ As String, D As String
    L = vbLf                 ' in-cell line break
    A = ChrW$(8217)          ' right single quote (customer's)
    LQ = ChrW$(8220)         ' left double quote
    RQ = ChrW$(8221)         ' right double quote
    D = ChrW$(8211)          ' en dash

    Dim t As String
    t = t & "ALERT REVIEW SUMMARY" & L
    t = t & "Alert ID: [Alert ID]" & L
    t = t & "Customer: [Subject Name]" & L
    t = t & "Program: [Program Name]" & L
    t = t & "Rule(s) Triggered: [Rule Name]" & L
    t = t & "Country Risk Rating: [Country Risk]" & L
    t = t & "________________________________________" & L
    t = t & "CUSTOMER PROFILE AND PROGRAM AFFILIATION" & L
    t = t & "[Subject Name] is a customer enrolled in the [Program Name] program. [Program Description]" & L
    t = t & L
    t = t & "The customer is located at [Location]." & L
    t = t & L
    t = t & "An external search of the address indicates that it is a <residential property / shared space / dedicated office / commercial property>." & L
    t = t & L
    t = t & "[External Research] An internet search for [Subject Name] returned results as <Individual: occupation and employer / Entity: line of business>." & L
    t = t & L
    t = t & "[Internal Research from Power BI]" & L
    t = t & "[Option 1 " & D & " Individual] Additionally, the customer" & A & "s full name and email address were identified through internal records as <" & LQ & "Name" & RQ & "> and <" & LQ
    t = t & "Email" & RQ
    t = t & ">, which <analysis of comparison of customer details between internal and external research, including support for validation of any information identified in the external research>." & L
    t = t & "[Option 2 " & D & " Entity] Additionally, the customer" & A & "s full name, address and industry were identified through internal records as <" & LQ & "Name" & RQ & ">, <" & LQ
    t = t & "Address" & RQ & "> and <" & LQ & "Industry type" & RQ
    t = t & ">, which <analysis of comparison, including any required additional OSDD in case of discrepancy, of customer details between internal and external research, including support for validation of any information identified in the external research>." & L
    t = t & L
    t = t & "[Negative News] A negative news search did not return any adverse or derogatory information about [Subject Name]." & L
    t = t & "________________________________________" & L
    t = t & "ALERT TRIGGER AND TRANSACTION REVIEW" & L
    t = t & "[Alert ID] alerted due to [Rule Name]." & L
    t = t & L
    t = t & "A review of the alerted transactions was conducted covering the period from [Start Date] to [End Date]. The review yielded the following results:" & L
    t = t & L
    t = t & "Between [Start Date] and [End Date], there were [Num Trans] transactions totaling $[Total Amount]. This includes [Num Dr] debit transactions of <transaction type(s): e.g., wires, ATM withdrawals, ACH transfers, card purchases etc.> totaling $[Dr] and [Num Cr] credit transactions of <transaction type(s): e.g., incoming wires, ACH credits, deposits etc.> totaling $[Cr]." & L
    t = t & L
    t = t & "[GALILEO ACCOUNT " & D
    t = t & " if applicable] These transactions are identified as <transaction type(s): e.g., wires, ATM withdrawals etc.> by <Customer Name / Third Party etc.> through their [Program Name] account <GAL Account #>." & L
    t = t & L
    t = t & "These transactions were conducted by [Subject Name] through their [Program Name] account [Account Numbers] with the following counterparties: [Counterparties]." & L
    t = t & L
    t = t & "[Subject Name] conducted transactions with the identified counterparty(s) in connection with <Activity Description " & D & " edit case to case specific>." & L
    t = t & L
    t = t & "[Option 1 " & D & " Individual] The alerted activity is consistent with the customer" & A
    t = t & "s expected financial behavior and does not appear to be inconsistent with the stated purpose of the account." & L
    t = t & "[Option 2 " & D & " Entity] <Include rationale for why the alerted activity is consistent with the customer" & A & "s profile and product usage " & D
    t = t & " consider the illustrative rationales from the comment, making it specific to the alert / transactions under review.>" & L
    t = t & L
    t = t & "[Non-alerted Activity Review]" & L
    t = t & "Other Activity Review:" & L
    t = t & "[Option 1 " & D
    t = t & " Individual] The non-alerted activity includes <# number of DR transactions> debit transactions of <transaction type(s): e.g., wires, ATM withdrawals, ACH transfers, card purchases etc.> totaling <$ DR amount> and <# number of CR transactions> credit transactions of <transaction type(s): e.g., incoming wires, ACH credits, deposits etc.> totaling <$ CR amount>." & L
    t = t & "[Option 2 " & D
    t = t & " Entity] The non-alerted activity includes debit transactions of <transaction type(s): e.g., wires, ATM withdrawals, ACH transfers, card purchases etc.> and credit transactions of <transaction type(s): e.g., incoming wires, ACH credits, deposits etc.>." & L
    t = t & L
    t = t & "<Highlight counterparties / transaction patterns from the non-alerted activity which are similar to / same as the counterparties / transaction patterns from the alerted activity and include rationale for their consistency.>" & L
    t = t & L
    t = t & "The other activity review appears reasonable and consistent with the alerted activity." & L
    t = t & "________________________________________" & L
    t = t & "[Counterparty Research]" & L
    t = t & "The following counterparty(s) were selected based on the identified customer rule types and their transaction value observed in the alerted transaction activity:" & L
    t = t & L
    t = t & "Counterparty 1: [CP 1]" & L
    t = t & "<Counterparty Details>" & L
    t = t & "[Negative News] A negative news search did not return any adverse or derogatory information about [CP 1]." & L
    t = t & L
    t = t & "Counterparty 2: [CP 2]" & L
    t = t & "<Counterparty Details>" & L
    t = t & "[Negative News] A negative news search did not return any adverse or derogatory information about [CP 2]." & L
    t = t & L
    t = t & "<Repeat for each counterparty reviewed; delete unused counterparty blocks.>" & L
    t = t & L
    t = t & "Additionally, <Counterparty name> was also identified as one of the counterparties, which is a listed entity. Hence, no due diligence was performed." & L
    t = t & L
    t = t & "[Option 1 " & D
    t = t & " Individual] Research identified <counterparty/customer> as <relationship/business/personal connection, if known>. No adverse media, sanctions, negative news, or other derogatory information was identified. The counterparties appear to have a reasonable relationship to the customer and are consistent with the customer"
    t = t & A & "s known profile and account usage." & L
    t = t & "[Option 2 " & D & " Entity] Research identified the counterparties as <customers/vendors/business partners/service providers> that appear reasonably connected to the customer" & A
    t = t & "s business operations. No adverse media, sanctions concerns, or other negative information were identified through available research." & L
    t = t & "________________________________________" & L
    t = t & "CONCLUSION" & L
    t = t & "[Option 1 " & D & " Individual] The customer is identified as <occupation/employment status> residing in <Country>. External research supports the customer" & A
    t = t & "s stated profile and identified no adverse information." & L
    t = t & "[Option 2 " & D
    t = t & " Entity] The customer is operating in <Country> with a stated business purpose of <line of business>. External research validates the existence of the business and is consistent with <customer"
    t = t & A & "s industry type>." & L
    t = t & L
    t = t & "<Document why the transaction pattern appears to be reasonable, including strong mitigating rationale to justify the same: key considerations " & D
    t = t & " location of customer, transaction location, customer occupation, relationship with counterparty etc.>" & L
    t = t & "<Document why non-alerted activity is reasonable, including strong rationale to justify the same: key considerations " & D
    t = t & " highlight any transaction pattern supporting the mitigating rationale. Use context from non-alerted activity to disposition the alert.>" & L
    t = t & "<Include commentary on transaction value, volume and frequency that supports the rationale.>" & L
    t = t & "<Include analysis of why the alerted activity is not suspicious. Describe the observed pattern / additional red flag(s) (if present) and the mitigation rationale against each of the rules the alert was triggered for.>" & L
    t = t & L
    t = t & "There is no evidence of structuring, layering, funnel account activity, or misuse of third-party accounts. No adverse media or negative information was identified for the customer or the business." & L
    t = t & "Currently, there are no red flags indicative of potential illicit or suspicious activity. Accordingly, it is recommended that this alert be closed with no further review warranted."

    On Error Resume Next
    wsT.Unprotect Password:="p7ss"
    ws7.Unprotect Password:="p7ss"
    On Error GoTo 0

    ' --- keep the text being replaced, once ---
    oldText = CStr(wsT.Cells(keyRow, "B").Value)
    If oldText <> t Then
        lastRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row
        wsT.Cells(lastRow + 1, "A").Value = "Non-Escalation_Closure (previous)"
        wsT.Cells(lastRow + 1, "B").Value = oldText
    End If

    wsT.Cells(keyRow, "B").Value = t

    ' --- [Country Risk] tag on Sheet7 (added once, then kept up to date) ---
    lastRow = ws7.Cells(ws7.Rows.Count, "J").End(xlUp).Row
    For r = 1 To lastRow
        If Trim$(CStr(ws7.Cells(r, "J").Value)) = "[Country Risk]" Then tagRow = r: Exit For
    Next r
    If tagRow = 0 Then tagRow = lastRow + 1
    ws7.Cells(tagRow, "J").Value = "[Country Risk]"
    ws7.Cells(tagRow, "K").Formula = _
        "=IF(Sheet1!$T$27="""","""",Sheet1!$T$27&IF(Sheet1!$G$27="""",""""," (""&Sheet1!$G$27&"")""))"

    MsgBox "Non-Escalation template (Individual / Entity) written to Template_DB!B" & keyRow & "." & vbCrLf & _
           "[Country Risk] tag added to Sheet7 row " & tagRow & "." & vbCrLf & vbCrLf & _
           "The previous text is kept on Template_DB as ""Non-Escalation_Closure (previous)""." & _
           IIf(keyRow <> 4, vbCrLf & vbCrLf & "Note: Sheet7's [MASTER_SHELL] formula points at Template_DB!B4, " & _
               "but this template is on row " & keyRow & ".", ""), vbInformation
End Sub
