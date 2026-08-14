Attribute VB_Name = "Module8"
Sub Start_Button_Create_Folders()
Dim wsHome As Worksheet
Dim ecmID As String
Dim desktopPath As String, mainFolderPath As String
Dim subFolderPath As String
Dim FSO As Object
Dim userProfile As String

Set wsHome = ThisWorkbook.Sheets("Sheet1")
ecmID = Trim(wsHome.Range("J9").Value)

If ecmID = "" Then
MsgBox "Action Denied: ECM ID is missing." & vbCrLf & "Please fill in the ECM ID before clicking START.", vbCritical, "Missing ID"
Exit Sub
End If

Set FSO = CreateObject("Scripting.FileSystemObject")
userProfile = Environ("USERPROFILE")

' STRICT PATHING MATCHING THE CONSOLIDATION MACRO
If FSO.FolderExists(userProfile & "\OneDrive - Community Federal Savings Bank\Desktop") Then
desktopPath = userProfile & "\OneDrive - Community Federal Savings Bank\Desktop"
Else
desktopPath = userProfile & "\Desktop"
End If

mainFolderPath = desktopPath & "\" & ecmID
subFolderPath = mainFolderPath & "\Transaction Files"

' Safely create Main Folder if it doesn't exist
If Not FSO.FolderExists(mainFolderPath) Then
FSO.CreateFolder mainFolderPath
End If

' Safely create the Transaction Files sub-folder if it doesn't exist.
' (Non Alerted / Lookback folders are no longer created here.)
If Not FSO.FolderExists(subFolderPath) Then
FSO.CreateFolder subFolderPath
End If

' Centralized audit ledger row - also seeds this case's own
' Register tab inside its Desktop\{ecmID}\{ecmID}_Audit_Log.xlsx.
modAuditLog.LogAuditEvent ecmID:=ecmID, _
AlertID:=Trim(wsHome.Range("J10").Value), _
customerName:=Trim(wsHome.Range("J13").Value), _
counterparties:=modAuditLog.GetCounterpartyList(wsHome), _
eventType:="Case Folder Created", _
outputFile:=mainFolderPath, _
toolVersion:="n/a", _
notes:="Transaction Files folder created"

' Automatically open the exact folder for the analyst
Shell "explorer.exe """ & subFolderPath & """", vbNormalFocus
End Sub

