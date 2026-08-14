Attribute VB_Name = "Module10"
Sub Trigger_PAD_Merge_Flow()
Dim padRunURL As String
Dim wsSettings As Worksheet, wsHome As Worksheet
Dim ECM As String, alert As String, entity As String
Dim jsonArgs As String
Dim entityChoice As String

' ==========================================
' 1. LOCATE SHEETS & GET URL
' ==========================================
On Error Resume Next
Set wsSettings = ThisWorkbook.Sheets("Backend_Settings")
Set wsHome = ThisWorkbook.Sheets("Sheet1")
On Error GoTo 0

If wsSettings Is Nothing Or wsHome Is Nothing Then
    MsgBox "CRITICAL ERROR: Could not find the required sheets!", vbCritical, "Error"
    Exit Sub
End If

padRunURL = Trim(wsSettings.Range("B4").Value)

If padRunURL = "" Then
    MsgBox "CRITICAL ERROR: Cell B4 is completely blank.", vbExclamation, "Error"
    Exit Sub
End If

' ==========================================
' 2. ENTITY SELECTOR
' ==========================================
entityChoice = InputBox("Which entity are you merging files for?" & vbCrLf & vbCrLf & _
                        "0 = Customer" & vbCrLf & _
                        "1 = Counterparty 1" & vbCrLf & _
                        "2 = Counterparty 2" & vbCrLf & _
                        "3 = Counterparty 3" & vbCrLf & _
                        "4 = Counterparty 4" & vbCrLf & _
                        "5 = Counterparty 5" & vbCrLf & _
                        "6 = Counterparty 6", "Select Entity", "0")
                        
If StrPtr(entityChoice) = 0 Or entityChoice = "" Then Exit Sub

' ==========================================
' 3. PULL DATA BASED ON EXACT SHEET LAYOUT
' ==========================================
Select Case Trim(entityChoice)
Case "0": entity = Trim(wsHome.Range("J13").Value) ' Customer (No changes)
Case "1": entity = "CP1_" & Trim(wsHome.Range("J18").Value) ' Counterparty 1
Case "2": entity = "CP2_" & Trim(wsHome.Range("J19").Value) ' Counterparty 2
Case "3": entity = "CP3_" & Trim(wsHome.Range("J20").Value) ' Counterparty 3
Case "4": entity = "CP4_" & Trim(wsHome.Range("J21").Value) ' Counterparty 4
Case "5": entity = "CP5_" & Trim(wsHome.Range("J22").Value) ' Counterparty 5
Case "6": entity = "CP6_" & Trim(wsHome.Range("J23").Value) ' Counterparty 6
Case Else
MsgBox "Invalid selection. Please run again and enter a number from 0 to 6.", vbExclamation, "Error"
Exit Sub
End Select

If entity = "" Or entity = "CP1_" Or entity = "CP2_" Or entity = "CP3_" Or entity = "CP4_" Or entity = "CP5_" Or entity = "CP6_" Then
MsgBox "The cell for the entity you selected is completely empty on Sheet1! Please check your data.", vbExclamation, "Missing Name"
Exit Sub
End If

' Grab the ECM and Alert IDs
ECM = Trim(wsHome.Range("J9").Value)
alert = Trim(wsHome.Range("J10").Value)

' ==========================================
' 4. SANITIZE DATA FOR POWER AUTOMATE
' ==========================================
ECM = Replace(Replace(Replace(Replace(ECM, "\", "\\"), """", "\"""), vbCrLf, " "), vbLf, " ")
alert = Replace(Replace(Replace(Replace(alert, "\", "\\"), """", "\"""), vbCrLf, " "), vbLf, " ")
entity = Replace(Replace(Replace(Replace(entity, "\", "\\"), """", "\"""), vbCrLf, " "), vbLf, " ")

jsonArgs = "{""ECMCase"":""" & ECM & """,""AlertID"":""" & alert & """,""EntityName"":""" & entity & """}"

' Safely URL-Encode the JSON
jsonArgs = Replace(jsonArgs, "%", "%25") ' Encode % first
jsonArgs = Replace(jsonArgs, """", "%22")
jsonArgs = Replace(jsonArgs, "{", "%7B")
jsonArgs = Replace(jsonArgs, "}", "%7D")
jsonArgs = Replace(jsonArgs, " ", "%20")
jsonArgs = Replace(jsonArgs, "&", "%26")
jsonArgs = Replace(jsonArgs, ":", "%3A")
jsonArgs = Replace(jsonArgs, ",", "%2C")
jsonArgs = Replace(jsonArgs, "+", "%2B")
jsonArgs = Replace(jsonArgs, "\", "%5C")
jsonArgs = Replace(jsonArgs, "/", "%2F")

padRunURL = padRunURL & "&inputArguments=" & jsonArgs

' ==========================================
' 5. THE SHELLEXECUTE BYPASS / FOLLOWHYPERLINK
' ==========================================
' This acts exactly like clicking a native link, protecting the quotes
On Error Resume Next
#If Mac Then
    ThisWorkbook.FollowHyperlink Address:=padRunURL
#Else
    CreateObject("Shell.Application").ShellExecute padRunURL
#End If

' Fallback if ShellExecute fails on Windows or if FollowHyperlink fails
If Err.Number <> 0 Then
    Err.Clear
    ThisWorkbook.FollowHyperlink Address:=padRunURL
End If

If Err.Number <> 0 Then
    MsgBox "Failed to launch Power Automate. Error: " & Err.Description, vbCritical, "Launch Error"
    Err.Clear
End If
On Error GoTo 0
End Sub
