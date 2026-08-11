Sub Run_Mass_Rename_v3()
' ========================================================
' ESC KEY HANDLER: Tell Excel not to show the default debug box
' ========================================================
Application.EnableCancelKey = xlErrorHandler
On Error GoTo CancelHandler

Dim ws As Worksheet: Set ws = ActiveSheet

Dim ECM As String: ECM = ws.Range("J10").Value
Dim sheetAlertID As String: sheetAlertID = ws.Range("J11").Value
Dim CustName As String: CustName = ws.Range("J14").Value

If ECM = "" Or sheetAlertID = "" Then MsgBox "Fill ECM (J10) & Alert ID (J11)!", vbCritical: Exit Sub

Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)

' --- FIX: SAFELY RESET THE DIALOG BOX ---
On Error Resume Next ' If Excel blocks filter changes, ignore the panic
With fd
    .title = "Select files to rename (Excel, PDF, Images, etc.)"
    .Filters.Clear
    .Filters.Add "All Files", "*.*", 1
    .AllowMultiSelect = True
End With
On Error GoTo CancelHandler ' Turn your normal error handler back on!
' ----------------------------------------
If fd.Show = -1 Then
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim fileItem As Variant
    Dim successCount As Integer: successCount = 0

    For Each fileItem In fd.SelectedItems
        Dim ext As String: ext = "." & FSO.GetExtensionName(fileItem)
        Dim fileName As String: fileName = FSO.GetFileName(fileItem)
        Dim newName As String

        ' --- 1. EXCEL LOGIC ---
        If LCase(ext) Like "*.xls*" Then
            ' Ask which kind of Excel transaction file this is.
            Dim xlType As String
            xlType = InputBox("Renaming: " & fileName & vbNewLine & vbNewLine & _
                              "What type of transaction file?" & vbNewLine & _
                              "1: Alerted Transactions" & vbNewLine & _
                              "2: Lookback Transactions" & vbNewLine & _
                              "3: Non Alerted Transactions" & vbNewLine & _
                              "4: Galileo Transaction" & vbNewLine & _
                              "5: CTS Report (keep file name)", "Excel Rename")
            If StrPtr(xlType) = 0 Then GoTo CancelHandler    ' Cancel
            If Trim(xlType) = "" Then GoTo NextFile           ' blank -> skip this file

            Select Case Trim(xlType)
                Case "1"   ' Alerted Transactions (existing behaviour)
                    Dim subAlert As String
                    subAlert = InputBox("Renaming: " & fileName & vbNewLine & vbNewLine & _
                                        "Enter Sub-Alert ID (or leave blank):", "Alerted Transactions")
                    If StrPtr(subAlert) = 0 Then GoTo CancelHandler
                    If subAlert <> "" Then
                        newName = ECM & "_" & sheetAlertID & "_Alerted Transactions(" & subAlert & ")"
                    Else
                        newName = ECM & "_" & sheetAlertID & "_Alerted Transactions"
                    End If

                Case "3"   ' Non Alerted Transactions - fixed name
                    newName = ECM & "_" & sheetAlertID & "_Non Alerted Transactions"

                Case "4"   ' Galileo Transaction - fixed name (moved here from PDF categories)
                    newName = ECM & "_" & sheetAlertID & "_Galileo Transaction"

                Case "5"   ' CTS Report - keep the file's own static name, just prefix it
                    Dim ctsBase As String
                    If Len(ext) > 0 And Len(fileName) >= Len(ext) And _
                       LCase(Right(fileName, Len(ext))) = LCase(ext) Then
                        ctsBase = Left(fileName, Len(fileName) - Len(ext))
                    Else
                        ctsBase = fileName
                    End If
                    newName = ECM & "_" & sheetAlertID & "_CTS Report_" & ctsBase

                Case "2"   ' Lookback Transactions (new) - date range in the name
                    Dim lbStart As String, lbEnd As String
                    lbStart = InputBox("Renaming: " & fileName & vbNewLine & vbNewLine & _
                                       "Enter LOOKBACK START date (mm.dd.yyyy):", "Lookback Start Date")
                    If StrPtr(lbStart) = 0 Then GoTo CancelHandler
                    If Trim(lbStart) = "" Then GoTo NextFile
                    lbEnd = InputBox("Renaming: " & fileName & vbNewLine & vbNewLine & _
                                     "Enter LOOKBACK END date (mm.dd.yyyy):", "Lookback End Date")
                    If StrPtr(lbEnd) = 0 Then GoTo CancelHandler
                    If Trim(lbEnd) = "" Then GoTo NextFile
                    newName = ECM & "_" & sheetAlertID & "_Lookback Transactions (" & _
                              FormatLookbackDate(lbStart) & " to " & FormatLookbackDate(lbEnd) & ")"

                Case Else
                    GoTo NextFile
            End Select

        ' --- 2. PDF/DOC/IMAGE LOGIC ---
        Else
            ' STEP 1: Select the Entity (Customer or Counterparty)
            Dim entityMenu As String
            Dim i As Integer, tempCP As String
            
            ' Build the dynamic menu with actual names
            entityMenu = "Who is this file for?" & vbNewLine & vbNewLine & "0: Customer (" & CustName & ")"
            For i = 1 To 6
                tempCP = ws.Range("J19").Offset(i - 1, 0).Value
                If tempCP <> "" Then
                    entityMenu = entityMenu & vbNewLine & i & ": Counterparty " & i & " (" & tempCP & ")"
                Else
                    entityMenu = entityMenu & vbNewLine & i & ": Counterparty " & i
                End If
            Next i

            Dim entityChoice As String
            entityChoice = InputBox("Renaming: " & fileName & vbNewLine & vbNewLine & entityMenu, "Step 1: Select Entity")

            If StrPtr(entityChoice) = 0 Then GoTo CancelHandler ' Clicked Cancel
            If entityChoice = "" Then GoTo NextFile ' Left blank and hit enter

            Dim identifier As String
            Dim docType As String: docType = ""

            ' STEP 2A: CUSTOMER LOGIC
            If entityChoice = "0" Then
                identifier = CustName
                
                Dim custChoice As String
                custChoice = InputBox("Category for " & CustName & ":" & vbNewLine & vbNewLine & _
                                  "1: External Search 1" & vbNewLine & _
                                  "2: External Search 2" & vbNewLine & _
                                  "3: External Search 3" & vbNewLine & _
                                  "4: Website" & vbNewLine & _
                                  "5: Google Translate" & vbNewLine & _
                                  "6: Customer Name + SSN" & vbNewLine & _
                                  "7: CTS Report_Customer KYC" & vbNewLine & _
                                  "8: Alert Write Up" & vbNewLine & _
                                  "9: Galileo Profile", "Step 2: Customer Category")

                If StrPtr(custChoice) = 0 Then GoTo CancelHandler
                If custChoice = "" Then GoTo NextFile

                Select Case custChoice
                    Case "1": docType = "External Search 1"
                    Case "2": docType = "External Search 2"
                    Case "3": docType = "External Search 3"
                    Case "4": docType = "Website"
                    Case "5": docType = "Google Translate"
                    Case "6": docType = "Customer Name + SSN"
                    Case "7": docType = "CTS Report_Customer KYC"
                    Case "8": docType = "Alert Write Up"
                    Case "9": docType = "Galileo Profile"
                    Case Else: GoTo NextFile
                End Select

            ' STEP 2B: COUNTERPARTY LOGIC
            ElseIf Val(entityChoice) >= 1 And Val(entityChoice) <= 6 Then
                Dim cpIdx As Integer: cpIdx = Val(entityChoice)
                Dim cpCellVal As String: cpCellVal = ws.Range("J19").Offset(cpIdx - 1, 0).Value

                If cpCellVal = "" Then
                    identifier = "Counterparty " & cpIdx
                Else
                    identifier = "Counterparty " & cpIdx & "_" & cpCellVal
                End If
                
                Dim cpChoice As String
                cpChoice = InputBox("Category for " & identifier & ":" & vbNewLine & vbNewLine & _
                                     "1: External Search 1" & vbNewLine & _
                                     "2: External Search 2" & vbNewLine & _
                                     "3: External Search 3" & vbNewLine & _
                                     "4: Website" & vbNewLine & _
                                     "5: Google Translate", "Step 2: CP Category")
                                     
                If StrPtr(cpChoice) = 0 Then GoTo CancelHandler
                If cpChoice = "" Then GoTo NextFile
                
                Select Case cpChoice
                    Case "1": docType = "External Search 1"
                    Case "2": docType = "External Search 2"
                    Case "3": docType = "External Search 3"
                    Case "4": docType = "Website"
                    Case "5": docType = "Google Translate"
                    Case Else: GoTo NextFile
                End Select
            Else
                GoTo NextFile
            End If

            ' Build the final file name
            If docType = "CTS Report_Customer KYC" Then
                ' Standalone label - no entity identifier in the name.
                newName = ECM & "_" & sheetAlertID & "_" & docType
            ElseIf docType <> "" Then
                newName = ECM & "_" & sheetAlertID & "_" & identifier & "_" & docType
            Else
                newName = ECM & "_" & sheetAlertID & "_" & identifier
            End If
        End If

        ' Clean the filename of any illegal characters before renaming
        newName = cleanFileName(newName)

        ' Rename the file
        On Error Resume Next
        Name fileItem As FSO.GetParentFolderName(fileItem) & "\" & newName & ext
        If Err.Number = 0 Then successCount = successCount + 1
        On Error GoTo CancelHandler ' Turn the ESC key trap back on
NextFile:
Next fileItem

    If successCount > 0 Then
        MsgBox successCount & " files have been renamed successfully!", vbInformation, "Rename Complete"
    End If
End If

' Reset normal Excel behavior when finished
Application.EnableCancelKey = xlInterrupt
Exit Sub
' ========================================================
' ESC KEY & CANCEL TRIGGER
' ========================================================
CancelHandler:
' Restore normal Excel behavior
Application.EnableCancelKey = xlInterrupt

If Err.Number = 18 Then
    MsgBox "Process Safely Cancelled." & vbCrLf & vbCrLf & "You pressed ESC to stop the macro.", vbInformation, "Aborted"
ElseIf Err.Number = 0 Then
    ' Triggered by hitting "Cancel" on an InputBox
    MsgBox "Renaming process aborted by user.", vbInformation, "Cancelled"
Else
    MsgBox "An unexpected error occurred:" & vbCrLf & Err.Description, vbCritical, "Error " & Err.Number
End If
End Sub

' --- HELPER FUNCTION TO REMOVE SPECIAL CHARACTERS ---
Function cleanFileName(ByVal strName As String) As String
Dim invalidChars As Variant
Dim i As Integer

' Array of characters that Windows does not allow in file names, plus line breaks
invalidChars = Array("<", ">", ":", """", "/", "\", "|", "?", "*", vbCr, vbLf, vbTab)

' Loop through and replace illegal characters with a space
For i = LBound(invalidChars) To UBound(invalidChars)
    strName = Replace(strName, invalidChars(i), " ")
Next i

' Clean up double spaces if any were created
Do While InStr(strName, "  ") > 0
    strName = Replace(strName, "  ", " ")
Loop

cleanFileName = Trim(strName)
End Function

' Formats a user-typed date as mm.dd.yyyy for the Lookback file name.
' If the entry can't be parsed as a date (e.g. they already typed it
' as 08.11.2026), the trimmed raw text is used as-is - so it always
' produces something sensible.
Function FormatLookbackDate(ByVal s As String) As String
On Error GoTo Fallback
FormatLookbackDate = Format$(CDate(Trim(s)), "mm.dd.yyyy")
Exit Function
Fallback:
FormatLookbackDate = Trim(s)
End Function

' ========================================================
' Run_Mass_Rename_v4 - batch rename, one form, all files at once.
'
' Same file-picker as v3, but instead of walking the analyst
' through a blocking InputBox per file, this hands the whole
' selection to UserForm2, which builds one live row (Entity/
' Category dropdowns + preview) per file and renames everything
' in one shot when "Rename All" is clicked. v3 is left in place
' untouched in case anything is still wired to it.
' ========================================================
Sub Run_Mass_Rename_v4()
Dim ws As Worksheet: Set ws = ActiveSheet

Dim ECM As String: ECM = ws.Range("J10").Value
Dim AlertID As String: AlertID = ws.Range("J11").Value
Dim CustName As String: CustName = ws.Range("J14").Value

If ECM = "" Or AlertID = "" Then MsgBox "Fill ECM (J10) & Alert ID (J11)!", vbCritical: Exit Sub

Dim CPNames(1 To 6) As String
Dim i As Long
For i = 1 To 6
CPNames(i) = ws.Range("J19").Offset(i - 1, 0).Value
Next i

Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
On Error Resume Next
With fd
.title = "Select files to rename (Excel, PDF, Images, etc.)"
.Filters.Clear
.Filters.Add "All Files", "*.*", 1
.AllowMultiSelect = True
End With
On Error GoTo 0

If fd.Show <> -1 Then Exit Sub   ' user cancelled the picker

UserForm2.InitRows fd.SelectedItems, ECM, AlertID, CustName, CPNames
UserForm2.Show vbModal
Unload UserForm2
End Sub

