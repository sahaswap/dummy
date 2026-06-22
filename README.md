'==================================================================
' OSINT Automation Tool
'
' Runs a batch of Google OSINT searches for a customer and their
' counterparties, saves each result page as a PDF straight into
' the ECM case folder, and writes an audit trail back to the
' workbook and the shared SharePoint tracker.
'
' Headless Edge does the heavy lifting. A small worker pool lets
' a couple of searches run in parallel so a typical 20-30 search
' job finishes in a few minutes instead of the better part of ten.
'
' To run: assign "SearchAndSavePDF_Direct" to your macro button.
'==================================================================

Option Explicit

' Direct Win32 imports - we reach past the VBA wrappers for things
' like foreground windows and modal dialogs because the wrappers
' don't behave well on locked-down / VDI hosts.
#If VBA7 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As LongPtr
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" _
    (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" _
    (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function MessageBoxW Lib "user32" ( _
    ByVal hwnd As Long, ByVal lpText As Long, _
    ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

' Win32 MessageBox flags used by TopMostMsgBox.
Private Const MB_OK           As Long = &H0
Private Const MB_YESNO        As Long = &H4
Private Const MB_ICONERROR    As Long = &H10
Private Const MB_ICONQUESTION As Long = &H20
Private Const MB_ICONWARNING  As Long = &H30
Private Const MB_ICONINFO     As Long = &H40
Private Const MB_TOPMOST      As Long = &H40000
Private Const MB_SETFOREGROUND As Long = &H10000
Private Const MB_SYSTEMMODAL  As Long = &H1000
Private Const IDOK  As Long = 1
Private Const IDYES As Long = 6
Private Const IDNO  As Long = 7

' WM_CLOSE for the handle-targeted browser-window close fallback.
Private Const WM_CLOSE As Long = &H10

' ---- Tunables ----
' Anything above MIN_PDF_SIZE_BYTES is treated as a real PDF.
' Anything below CAPTCHA_SIZE_HINT is suspiciously small and flagged
' for a retry, since Google's interstitial/CAPTCHA pages tend to
' render tiny compared to a normal 100-result SERP.
Private Const TOOL_VERSION As String = "2.5.12"
Private Const MIN_PDF_SIZE_BYTES As Long = 5120
Private Const CAPTCHA_SIZE_HINT As Long = 80000

Private USE_HEADLESS As Boolean
Private Const TIMING_PROFILE As String = "FAST"         ' FAST | NORMAL | SLOW

' Parallel dispatch. Two workers is the sweet spot on a 2-vCPU VDI -
' a third worker just fights the others for CPU and pushes tasks
' past the timeout.
Private Const MAX_PARALLEL As Long = 2
Private Const LAUNCH_STAGGER_SEC As Single = 0.3        ' breathing room between launches
Private Const POLL_INTERVAL_SEC As Single = 0.3         ' how often we check for a finished PDF
Private Const TASK_TIMEOUT_SEC As Long = 75             ' kill a task that is clearly stuck
Private Const RETRY_TIMEOUT_SEC As Long = 90            ' give retries a little more rope
Private Const RETRY_COOLDOWN_SEC As Single = 5          ' unused since v2.5.10 (inline retries) - kept in case Google rate-limiting forces a cooldown back in

' Leave the drive alone if it is running out of space.
Private Const DISK_ABORT_GB As Double = 2
Private Const DISK_WARN_GB As Double = 5

' While the macro runs we briefly drop the priority of the usual CPU
' hogs (Teams, Outlook, OneDrive) so Edge gets a cleaner run at the
' 2 vCPUs. Flip to False if it ever causes trouble on a given VDI.
Private Const TUNE_BACKGROUND_PRIORITIES As Boolean = True

' Shared tracker that every analyst's run ends up appending to.
Private Const MASTER_PATH As String = _
    "https://cfsb-my.sharepoint.com/personal/msrivastava_cfsb_com/Documents/" & _
    "L1 Beta/Beta 2.4_Feedbacks & Issues Encountered.xlsx"
Private Const MASTER_FILENAME As String = "Beta 2.4_Feedbacks & Issues Encountered.xlsx"

' Module state. bAbort is flipped to True when the user presses ESC
' or any branch wants to unwind cleanly.
Private bAbort As Boolean

Private m_origCalc As XlCalculation
Private m_origScreenUpdate As Boolean
Private m_origEvents As Boolean
Private m_origAlerts As Boolean
Private m_origCursor As XlMousePointer
Private m_origPrintComm As Boolean
Private m_runtimeApplied As Boolean

' PIDs of processes we temporarily knocked down to BelowNormal, so
' we know which ones to lift back to Normal at the end of the run.
Private m_tunedPids As Collection

' Entry point. Wire this up to your macro button.
Sub SearchAndSavePDF_Direct()

    On Error GoTo ErrorHandler

    ' Start every run with a fresh debug log so the tail is always this session.
    On Error Resume Next
    Kill Environ("TEMP") & "\osint_debug.log"
    On Error GoTo ErrorHandler
    LogStep "=== Macro started (v" & TOOL_VERSION & ") ==="

    ' Self-heal entry guard. If the previous run died before its
    ' AutoRestoreRuntime fired (VBE Reset, "End" on a runtime error,
    ' Excel killed mid-batch), Excel is still in our muted mode and
    ' every formula on the workbook is showing the yellow "Stale"
    ' badge with the value struck through. Backend / hidden sheets
    ' can't be fixed by the analyst from the ribbon, so we recover
    ' here before doing anything else. No-op if state is already clean.
    ForceClearStaleOSINTState

    ' ---- Declarations ----
    Dim ws As Worksheet, auditWs As Worksheet, masterWs As Worksheet
    Dim WshShell As Object, FSO As Object, ghostApp As Object
    Dim masterWb As Workbook, wb As Workbook

    Dim i As Long, fileCounter As Long, totalCPs As Long
    Dim entityCount As Long, actualSearchCount As Long
    Dim cpLoopPosition As Long, hdrIdx As Long
    Dim nextAuditRow As Long, mRow As Long
    Dim wasAlreadyOpen As Boolean, headersOK As Boolean

    Dim ecmCase As String, alertID As String
    Dim custName As String, custAddr As String
    Dim custNameNegNews As String, custNameNegNewsNoMiddle As String
    Dim additionalCustAddr As String
    Dim cpName As String, cpAddr As String
    Dim cpNameNegNews(19 To 24) As String
    Dim cpNameNegNewsNoMiddle(19 To 24) As String

    Dim baseFileName As String, desktopPath As String
    Dim caseFolderPath As String, mainFolderPath As String
    Dim negWords As String
    Dim userProfile As String
    Dim custLabel As String, cpLabel As String

    Dim StartTime As Date
    Dim ElapsedSeconds As Long, Minutes As Long, Seconds As Long
    Dim timeString As String

    Dim nameParts() As String
    Dim addAddrResponse As VbMsgBoxResult

    Dim expectedHeaders As Variant, headerMsg As String
    Dim spSyncStatus As String      ' "OK" / "FAILED: ..." / "SKIPPED"
    Dim allTasks As Collection

    '--------------------------------------------------------------
    ' INITIALIZATION
    '--------------------------------------------------------------
    bAbort = False
    additionalCustAddr = ""
    custNameNegNewsNoMiddle = ""
    wasAlreadyOpen = False
    totalCPs = 0
    spSyncStatus = "SKIPPED"
    actualSearchCount = 0

    If Not SheetExists(ThisWorkbook, "Sheet1") Then
        MsgBox "Required sheet 'Sheet1' is missing. Please use the original template.", _
            vbCritical, "Template Error"
        Exit Sub
    End If
    Set ws = ThisWorkbook.Sheets("Sheet1")

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Figure out where this machine keeps its Desktop - it might be
    ' under the OneDrive-synced profile or the plain local one.
    ' We prefer OneDrive so the case folder syncs automatically.
    userProfile = Environ("USERPROFILE")
    If FSO.FolderExists(userProfile & "\OneDrive - Community Federal Savings Bank\Desktop") Then
        desktopPath = userProfile & "\OneDrive - Community Federal Savings Bank\Desktop"
    Else
        desktopPath = userProfile & "\Desktop"
    End If

    ecmCase = SanitizeFileNamePart(Trim(CStr(ws.Range("J10").Value)))
    alertID = SanitizeFileNamePart(Trim(CStr(ws.Range("J11").Value)))
    custName = SanitizeFileNamePart(Trim(CStr(ws.Range("J14").Value)))
    custAddr = SanitizeFileNamePart(Trim(CStr(ws.Range("J15").Value)))

    If custName = "" And Application.CountA(ws.Range("J19:J24")) = 0 Then
        MsgBox "I don't see a Customer Name or any Counterparties to search. " & _
            "Please fill them in first!", vbExclamation, "Nothing to Search"
        Exit Sub
    End If

    If ecmCase = "" Then
        MsgBox "I don't see an ECM Case ID in cell J10. Please enter it before running.", _
            vbExclamation, "Missing Case ID"
        Exit Sub
    End If

    negWords = " AND (arrest OR corruption OR sentencing OR money laundering OR AML " & _
        "OR launder OR embezzle OR evas OR evad OR crime OR corrupt OR bribe OR " & _
        "theft OR extort OR drug OR traffic OR trafficking OR felony OR sanction " & _
        "OR counterfeit OR terror)"

    baseFileName = ecmCase & "_" & alertID
    baseFileName = Replace(baseFileName, " ", "_")
    Do While InStr(baseFileName, "__") > 0
        baseFileName = Replace(baseFileName, "__", "_")
    Loop
    ' Trim any stray underscores the sanitiser left at the edges.
    Do While Len(baseFileName) > 0 And Right(baseFileName, 1) = "_"
        baseFileName = Left(baseFileName, Len(baseFileName) - 1)
    Loop
    Do While Len(baseFileName) > 0 And Left(baseFileName, 1) = "_"
        baseFileName = Mid(baseFileName, 2)
    Loop

    caseFolderPath = desktopPath & "\" & ecmCase
    If Not FSO.FolderExists(caseFolderPath) Then FSO.CreateFolder caseFolderPath

    mainFolderPath = caseFolderPath & "\OSDD Searches"
    If Not FSO.FolderExists(mainFolderPath) Then FSO.CreateFolder mainFolderPath

    totalCPs = Application.CountA(ws.Range("J19:J24"))

    ' Ask about Non-English names (Manual/Headless mode)
    LogStep "STEP 2: before non-english prompt"
    Application.StatusBar = "OSINT: Waiting for Non-English search decision..."
    Beep
    Dim nonEngResp As Long
    nonEngResp = TopMostMsgBox( _
        "Do you want to do searches for non-english names? " & vbCrLf & vbCrLf & _
        "(Choosing Yes will open browser windows visibly to allow for page translation. Choosing No will use the fast invisible searches.)", _
        "Non-English Searches", MB_YESNO Or MB_ICONQUESTION Or MB_TOPMOST)
    If nonEngResp = IDYES Then
        USE_HEADLESS = False
    Else
        USE_HEADLESS = True
    End If
    LogStep "STEP 2: non-english prompt returned USE_HEADLESS=" & CStr(USE_HEADLESS)

    ' Kick off DNS/TLS warmup early so the handshake is already
    ' done by the time the user is through the prompts.
    LogStep "STEP 2: before AutoPrewarmDNS"
    AutoPrewarmDNS
    LogStep "STEP 2: after AutoPrewarmDNS"

    ' Profile pre-warm: launch a tiny headless Edge against each
    ' worker's user-data-dir so the profile is populated and any
    ' first-run cost is paid while the analyst is still answering
    ' prompts. Fire-and-forget; DrainPrewarmAndCleanup picks up the
    ' pieces just before the real dispatch begins.
    LogStep "STEP 2: before PrewarmEdgeProfiles"
    PrewarmEdgeProfiles
    LogStep "STEP 2: after PrewarmEdgeProfiles"

    ' Collect the small bits of input we need from the analyst:
    ' negative-news variants and any extra addresses.
    If custName <> "" Then
        LogStep "STEP 3: before customer negnews InputBox"
        Application.StatusBar = "OSINT: Waiting for Customer Negative News input..."
        Beep
        ForcePromptToFront
        custNameNegNews = InputBox( _
            "Do you want to run a Negative News search for the Customer? " & _
            "You can edit the name below, or delete it completely to skip.", _
            "Customer Negative News", custName)
        LogStep "STEP 3: customer negnews returned, len=" & Len(custNameNegNews)
        If StrPtr(custNameNegNews) = 0 Then Exit Sub

        ' Collapse runs of spaces so "John  Smith" doesn't look like a middle-name case.
        custNameNegNews = NormalizeSpaces(custNameNegNews)
        LogStep "STEP 3: after NormalizeSpaces, len=" & Len(custNameNegNews)

        If Trim(custNameNegNews) <> "" Then
            nameParts = Split(Trim(custNameNegNews), " ")
            LogStep "STEP 3: Split into " & (UBound(nameParts) + 1) & " parts"
            If UBound(nameParts) >= 2 Then
                custNameNegNewsNoMiddle = nameParts(0) & " " & nameParts(UBound(nameParts))
                LogStep "STEP 3: before middle-name InputBox"
                Application.StatusBar = "OSINT: Waiting for Middle Name confirmation..."
                Beep
                ForcePromptToFront
                custNameNegNewsNoMiddle = InputBox( _
                    "I noticed the customer might have a middle name (" & custNameNegNews & ")." & _
                    vbCrLf & vbCrLf & "If you want to run an additional search WITHOUT the middle " & _
                    "name, you can edit/confirm the name below." & vbCrLf & _
                    "(Clear the text or click Cancel to skip this extra search)", _
                    "Middle Name Detected", custNameNegNewsNoMiddle)
                LogStep "STEP 3: middle-name InputBox returned"
                If StrPtr(custNameNegNewsNoMiddle) = 0 Then custNameNegNewsNoMiddle = ""
                custNameNegNewsNoMiddle = NormalizeSpaces(custNameNegNewsNoMiddle)
            End If
        End If

        LogStep "STEP 3: before Sigma address MsgBox"
        Application.StatusBar = "OSINT: Waiting for Sigma Address decision..."
        Dim sigmaResp As Long
        sigmaResp = TopMostMsgBox( _
            "Did you identify any additional addresses for the Customer via Sigma?", _
            "Sigma Address Check", MB_YESNO Or MB_ICONQUESTION Or MB_TOPMOST)
        LogStep "STEP 3: Sigma TopMostMsgBox returned " & sigmaResp
        If sigmaResp = IDYES Then
            LogStep "STEP 3: before additional-address InputBox"
            Beep
            ForcePromptToFront
            additionalCustAddr = InputBox( _
                "Please paste the additional address identified via Sigma below:", _
                "Enter Additional Address")
            LogStep "STEP 3: additional-address returned"
            If StrPtr(additionalCustAddr) = 0 Then Exit Sub
            additionalCustAddr = NormalizeSpaces(additionalCustAddr)
        End If
    End If
    LogStep "STEP 3: customer prompts complete"

    For i = 19 To 24
        cpName = SanitizeFileNamePart(Trim(CStr(ws.Range("J" & i).Value)))
        If cpName <> "" Then
            Application.StatusBar = "OSINT: Waiting for CP '" & cpName & "' Negative News input..."
            Beep
            ForcePromptToFront
            cpNameNegNews(i) = InputBox( _
                "Do you want to run a Negative News search for this Counterparty? " & _
                "You can edit the name or clear it to skip.", _
                "Counterparty: " & cpName, cpName)
            If StrPtr(cpNameNegNews(i)) = 0 Then Exit Sub
            cpNameNegNews(i) = NormalizeSpaces(cpNameNegNews(i))

            If Trim(cpNameNegNews(i)) <> "" Then
                nameParts = Split(Trim(cpNameNegNews(i)), " ")
                If UBound(nameParts) >= 2 Then
                    cpNameNegNewsNoMiddle(i) = nameParts(0) & " " & nameParts(UBound(nameParts))
                    Application.StatusBar = "OSINT: Waiting for CP '" & cpName & "' Middle Name confirmation..."
                    Beep
                    ForcePromptToFront
                    cpNameNegNewsNoMiddle(i) = InputBox( _
                        "I noticed the counterparty might have a middle name (" & _
                        cpNameNegNews(i) & ")." & vbCrLf & vbCrLf & _
                        "If you want to run an additional search WITHOUT the middle name, " & _
                        "you can edit/confirm the name below." & vbCrLf & _
                        "(Clear the text or click Cancel to skip this extra search)", _
                        "Middle Name Detected", cpNameNegNewsNoMiddle(i))
                    If StrPtr(cpNameNegNewsNoMiddle(i)) = 0 Then cpNameNegNewsNoMiddle(i) = ""
                    cpNameNegNewsNoMiddle(i) = NormalizeSpaces(cpNameNegNewsNoMiddle(i))
                End If
            End If
        End If
    Next i

    LogStep "STEP 3: before Starting Searches MsgBox"
    Application.StatusBar = "OSINT: Waiting for Starting-Searches confirmation..."
    TopMostMsgBox _
        "I'm going to run the searches now! Please keep your hands off the mouse " & _
        "and keyboard so I don't get interrupted." & vbCrLf & vbCrLf & _
        "(If things go wrong, just hold down the ESC key for 2 seconds to stop me).", _
        "Starting Searches", MB_OK Or MB_ICONINFO
    LogStep "STEP 3: Starting Searches MsgBox dismissed -> proceeding to searches"

    ' Bail out if the drive is dangerously full, then flip Excel into
    ' its fast-mode (calc manual, no screen update, etc.).
    If Not CheckDiskSpace() Then Exit Sub

    AutoPrepareRuntime
    StartTime = Now
    LogStep "SEARCH PHASE start - total entities to process: " & (IIf(custName <> "", 1, 0) + totalCPs)

    ' Build one flat queue for the customer + every counterparty.
    ' Keeping all tasks in a single pool means the workers never sit
    ' idle waiting for the next entity's batch to start.
    Set allTasks = New Collection
    Application.StatusBar = "OSINT: Queueing searches..."

    ' ---- Customer tasks ----
    If custName <> "" Then
        custLabel = "Customer (" & custName & ")"
        fileCounter = 1

        allTasks.Add Array(custName, _
            mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Google.pdf", _
            1, "Google name", custLabel)
        fileCounter = fileCounter + 1

        If custAddr <> "" Then
            allTasks.Add Array(custAddr, _
                mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Address.pdf", _
                1, "Address", custLabel)
            fileCounter = fileCounter + 1

            allTasks.Add Array(custName & " + " & custAddr, _
                mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Google+Address.pdf", _
                1, "Name + Address", custLabel)
            fileCounter = fileCounter + 1
        End If

        If Trim(additionalCustAddr) <> "" Then
            allTasks.Add Array(custName & " + " & additionalCustAddr, _
                mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Name+Additional Address.pdf", _
                1, "Name + Additional Address", custLabel)
            fileCounter = fileCounter + 1
        End If

        If Trim(custNameNegNews) <> "" Then
            allTasks.Add Array("""" & custNameNegNews & """" & negWords, _
                mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 1.pdf", _
                1, "Negative News page 1", custLabel)
            fileCounter = fileCounter + 1
            allTasks.Add Array("""" & custNameNegNews & """" & negWords, _
                mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 2.pdf", _
                2, "Negative News page 2", custLabel)
            fileCounter = fileCounter + 1
        End If

        If Trim(custNameNegNewsNoMiddle) <> "" And _
           StrComp(custNameNegNewsNoMiddle, custNameNegNews, vbTextCompare) <> 0 Then
            allTasks.Add Array("""" & custNameNegNewsNoMiddle & """" & negWords, _
                mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 3.pdf", _
                1, "Negative News w/o middle, page 1", custLabel)
            fileCounter = fileCounter + 1
            allTasks.Add Array("""" & custNameNegNewsNoMiddle & """" & negWords, _
                mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 4.pdf", _
                2, "Negative News w/o middle, page 2", custLabel)
            fileCounter = fileCounter + 1
        End If
    End If

    ' ---- Counterparty tasks (appended to same queue) ----
    cpLoopPosition = 0
    For i = 19 To 24
        cpName = SanitizeFileNamePart(Trim(CStr(ws.Range("J" & i).Value)))
        cpAddr = SanitizeFileNamePart(Trim(CStr(ws.Range("T" & i).Value)))

        If cpName <> "" Then
            cpLoopPosition = cpLoopPosition + 1
            cpLabel = "CP " & cpLoopPosition & " of " & totalCPs & " (" & cpName & ")"
            fileCounter = 1

            allTasks.Add Array(cpName, _
                mainFolderPath & "\" & baseFileName & "_CP" & (i - 18) & "_" & fileCounter & "_Google.pdf", _
                1, "Google name", cpLabel)
            fileCounter = fileCounter + 1

            If cpAddr <> "" Then
                allTasks.Add Array(cpAddr, _
                    mainFolderPath & "\" & baseFileName & "_CP" & (i - 18) & "_" & fileCounter & "_Address.pdf", _
                    1, "Address", cpLabel)
                fileCounter = fileCounter + 1

                allTasks.Add Array(cpName & " + " & cpAddr, _
                    mainFolderPath & "\" & baseFileName & "_CP" & (i - 18) & "_" & fileCounter & "_Google+Address.pdf", _
                    1, "Name + Address", cpLabel)
                fileCounter = fileCounter + 1
            End If

            If Trim(cpNameNegNews(i)) <> "" Then
                allTasks.Add Array("""" & cpNameNegNews(i) & """" & negWords, _
                    mainFolderPath & "\" & baseFileName & "_CP" & (i - 18) & "_" & fileCounter & "_Negative News 1.pdf", _
                    1, "Negative News page 1", cpLabel)
                fileCounter = fileCounter + 1
                allTasks.Add Array("""" & cpNameNegNews(i) & """" & negWords, _
                    mainFolderPath & "\" & baseFileName & "_CP" & (i - 18) & "_" & fileCounter & "_Negative News 2.pdf", _
                    2, "Negative News page 2", cpLabel)
                fileCounter = fileCounter + 1
            End If

            If Trim(cpNameNegNewsNoMiddle(i)) <> "" And _
               StrComp(cpNameNegNewsNoMiddle(i), cpNameNegNews(i), vbTextCompare) <> 0 Then
                allTasks.Add Array("""" & cpNameNegNewsNoMiddle(i) & """" & negWords, _
                    mainFolderPath & "\" & baseFileName & "_CP" & (i - 18) & "_" & fileCounter & "_Negative News 3.pdf", _
                    1, "Negative News w/o middle, page 1", cpLabel)
                fileCounter = fileCounter + 1
                allTasks.Add Array("""" & cpNameNegNewsNoMiddle(i) & """" & negWords, _
                    mainFolderPath & "\" & baseFileName & "_CP" & (i - 18) & "_" & fileCounter & "_Negative News 4.pdf", _
                    2, "Negative News w/o middle, page 2", cpLabel)
                fileCounter = fileCounter + 1
            End If
        End If
    Next i

    ' ---- Single dispatch across a shared worker pool ----
    If allTasks.Count > 0 Then
        If USE_HEADLESS Then
            Set allTasks = SortTasksHeaviestFirst(allTasks)
        End If

        Application.StatusBar = "OSINT: Queued " & allTasks.Count & " searches - dispatching..."
        LogStep "SEARCH PHASE - queued " & allTasks.Count & " tasks across all entities (heaviest first)"
        actualSearchCount = RunSearchBatch(allTasks, WshShell)
        If bAbort Then GoTo AbortProcess
    End If

    ' Elapsed time - using DateDiff instead of subtraction so a run
    ' that straddles midnight doesn't come out negative.
    ElapsedSeconds = DateDiff("s", StartTime, Now)
    Minutes = ElapsedSeconds \ 60
    Seconds = ElapsedSeconds Mod 60
    LogStep "SEARCH PHASE end - elapsed=" & ElapsedSeconds & "s, searches=" & actualSearchCount
    If Minutes > 0 Then
        timeString = Minutes & " minute(s) and " & Seconds & " second(s)"
    Else
        timeString = Seconds & " second(s)"
    End If

    entityCount = IIf(custName <> "" Or custNameNegNews <> "", 1, 0) + totalCPs
    ' actualSearchCount is the authoritative successful-PDF counter.
    ' We trust this over re-scanning the folder - OneDrive can lie
    ' about what's on disk while a sync is still in flight.

    ' Local audit log row.
    Set auditWs = GetOrCreateAuditSheet(ThisWorkbook)
    nextAuditRow = auditWs.Cells(auditWs.Rows.Count, "A").End(xlUp).Row + 1
    auditWs.Cells(nextAuditRow, 1).Value = Now
    auditWs.Cells(nextAuditRow, 2).Value = Environ("USERNAME")
    auditWs.Cells(nextAuditRow, 3).Value = ecmCase
    auditWs.Cells(nextAuditRow, 4).Value = entityCount
    auditWs.Cells(nextAuditRow, 5).Value = actualSearchCount
    auditWs.Cells(nextAuditRow, 6).Value = timeString
    auditWs.Cells(nextAuditRow, 7).Value = TOOL_VERSION
    ThisWorkbook.Save

    ' Hand Excel back to the user. Folder open is deferred until
    ' after they dismiss the success dialog (analyst-requested UX).
    AutoRestoreRuntime

    ' Push a row to the shared SharePoint tracker so the team can
    ' see this run. Failures here are non-fatal - the local audit
    ' log already has the authoritative record.
    Application.StatusBar = "OSINT: Pushing run to SharePoint master tracker..."

    For Each wb In Application.Workbooks
        If wb.Name = MASTER_FILENAME Then
            Set masterWb = wb
            wasAlreadyOpen = True
            Exit For
        End If
    Next wb

    If masterWb Is Nothing Then
        Set ghostApp = CreateObject("Excel.Application")
        ghostApp.Visible = False
        ghostApp.DisplayAlerts = False
        ghostApp.EnableEvents = False

        On Error Resume Next
        Err.Clear
        Set masterWb = ghostApp.Workbooks.Open(fileName:=MASTER_PATH, UpdateLinks:=False)
        If masterWb Is Nothing Then
            spSyncStatus = "FAILED: " & IIf(Err.Number <> 0, _
                "Err " & Err.Number & " - " & Err.Description, _
                "WebDAV open returned Nothing (offline, auth, or file renamed?)")
        End If
        On Error GoTo ErrorHandler
    End If

    If Not masterWb Is Nothing Then
        If Not masterWb.ReadOnly Then
            Set masterWs = masterWb.Sheets("Sheet1")
            expectedHeaders = Array("Date & Time", "Analyst ID", "ECM Case ID", _
                "Total Entities", "Total Searches", "Time Taken", "Tool Version")
            headersOK = True
            headerMsg = ""

            For hdrIdx = LBound(expectedHeaders) To UBound(expectedHeaders)
                If masterWs.Cells(1, hdrIdx + 1).Value <> expectedHeaders(hdrIdx) Then
                    headersOK = False
                    headerMsg = headerMsg & "- Col " & _
                        Split(masterWs.Cells(1, hdrIdx + 1).Address, "$")(1) & _
                        " expected '" & expectedHeaders(hdrIdx) & "' but found '" & _
                        masterWs.Cells(1, hdrIdx + 1).Value & "'" & vbCrLf
                End If
            Next hdrIdx

            If headersOK Then
                mRow = masterWs.Cells(masterWs.Rows.Count, "A").End(xlUp).Row + 1
                masterWs.Cells(mRow, 1).Value = Now
                masterWs.Cells(mRow, 2).Value = Environ("USERNAME")
                masterWs.Cells(mRow, 3).Value = ecmCase
                masterWs.Cells(mRow, 4).Value = entityCount
                masterWs.Cells(mRow, 5).Value = actualSearchCount
                masterWs.Cells(mRow, 6).Value = timeString
                masterWs.Cells(mRow, 7).Value = TOOL_VERSION

                If wasAlreadyOpen Then
                    masterWb.Save
                Else
                    masterWb.Close SaveChanges:=True
                End If
                spSyncStatus = "OK"
            Else
                spSyncStatus = "FAILED: header mismatch -" & vbCrLf & headerMsg
                If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
            End If
        Else
            spSyncStatus = "FAILED: master tracker is read-only (another user has it open?)"
            If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
        End If
    End If

    If Not ghostApp Is Nothing Then
        On Error Resume Next
        ghostApp.Quit
        Set ghostApp = Nothing
        On Error GoTo ErrorHandler
    End If

    Application.StatusBar = False

    '--------------------------------------------------------------
    ' STEP 11: Final summary dialog -> folder open on OK
    '--------------------------------------------------------------
    AppActivate Application.Caption

    Dim summary As String
    summary = "All done! I saved " & actualSearchCount & " PDF(s) directly into:" & vbCrLf & _
              "  " & ecmCase & "\OSDD Searches" & vbCrLf & vbCrLf & _
              "Total search time: " & timeString

    MsgBox summary, vbInformation, "Success"

    ' User dismissed the dialog - now pop the case folder so they
    ' land on the freshly-saved PDFs.
    On Error Resume Next
    Shell "explorer.exe """ & mainFolderPath & """", vbNormalFocus
    On Error GoTo ErrorHandler

    Set WshShell = Nothing
    Set FSO = Nothing
    GoTo CleanExit

'------------------------------------------------------------------
' ABORT PATH (user pressed ESC)
'------------------------------------------------------------------
AbortProcess:
    AutoRestoreRuntime
    On Error Resume Next
    AppActivate Application.Caption
    On Error GoTo 0
    MsgBox "You safely canceled the process. Any PDFs already downloaded " & _
           "are in the case folder." & vbCrLf & vbCrLf & _
           "Note: any Edge processes still running will finish on their own " & _
           "and their output will also land in the folder.", _
           vbInformation, "Canceled"
    GoTo CleanExit

' Catch-all for anything unexpected. We log the error on the audit
' sheet and show the analyst a dialog instead of dying silently.
ErrorHandler:
    Dim errN As Long, errD As String, errS As String
    errN = Err.Number
    errD = Err.Description
    errS = Err.Source

    On Error Resume Next
    LogStep "ERROR Err " & errN & ": " & errD & " [src: " & errS & "]"

    ' Log to audit sheet (col H for error diag). We intentionally do
    ' NOT call ThisWorkbook.Save here - if the workbook is on OneDrive
    ' and sync is blocked, Save can hang indefinitely and freeze Excel.
    ' Analyst can save manually when they dismiss the MsgBox.
    If auditWs Is Nothing Then Set auditWs = GetOrCreateAuditSheet(ThisWorkbook)
    Dim errRow As Long
    errRow = auditWs.Cells(auditWs.Rows.Count, "A").End(xlUp).Row + 1
    auditWs.Cells(errRow, 1).Value = Now
    auditWs.Cells(errRow, 2).Value = Environ("USERNAME")
    auditWs.Cells(errRow, 3).Value = ecmCase
    auditWs.Cells(errRow, 4).Value = entityCount
    auditWs.Cells(errRow, 5).Value = actualSearchCount
    auditWs.Cells(errRow, 6).Value = "ERROR"
    auditWs.Cells(errRow, 7).Value = TOOL_VERSION
    auditWs.Cells(errRow, 8).Value = "Err " & errN & ": " & errD & " [src: " & errS & "]"

    AutoRestoreRuntime
    Application.StatusBar = False

    MsgBox "Something went wrong and I had to stop." & vbCrLf & vbCrLf & _
           "Error " & errN & ": " & errD & vbCrLf & _
           "Source: " & errS & vbCrLf & vbCrLf & _
           "Details are in the Audit_Log sheet (column H) and in " & _
           Environ("TEMP") & "\osint_debug.log. " & _
           "Please Ctrl+S to save the workbook if you want to keep the audit row. " & _
           "Any PDFs already saved are safe in the case folder. " & _
           "You can re-run the macro; already-downloaded searches will be " & _
           "overwritten with fresh copies.", _
           vbCritical + vbSystemModal, "OSINT Error"
    On Error GoTo 0
    ' fall through to CleanExit

CleanExit:
    On Error Resume Next
    If Not ghostApp Is Nothing Then
        ghostApp.Quit
        Set ghostApp = Nothing
    End If
    AutoRestoreRuntime
    Application.StatusBar = False
    On Error GoTo 0

End Sub

' Refuses to run if the user's drive is close to full. Batch PDF
' runs can easily eat a few hundred MB and a half-written case
' folder is worse than no folder at all.
Private Function CheckDiskSpace() As Boolean
    On Error Resume Next
    Dim FSO As Object, drv As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set drv = FSO.GetDrive(FSO.GetDriveName(Environ("USERPROFILE")))

    Dim freeGB As Double
    freeGB = drv.FreeSpace / (1024 ^ 3)
    On Error GoTo 0

    If freeGB < DISK_ABORT_GB Then
        MsgBox "CRITICAL: Only " & Format(freeGB, "0.0") & " GB free on your VDI." & vbCrLf & _
               "The tool needs at least " & DISK_ABORT_GB & " GB to run safely. Aborting." & vbCrLf & vbCrLf & _
               "Please empty Recycle Bin, clear Downloads, or contact IT to free up space.", _
               vbCritical, "Disk Full"
        CheckDiskSpace = False
    ElseIf freeGB < DISK_WARN_GB Then
        Dim r As VbMsgBoxResult
        r = MsgBox("WARNING: Only " & Format(freeGB, "0.0") & " GB free on your VDI." & vbCrLf & _
                   "A typical 20-search run uses ~700 MB. Continue anyway?", _
                   vbYesNo + vbExclamation, "Low Disk Space")
        CheckDiskSpace = (r = vbYes)
    Else
        CheckDiskSpace = True
    End If
End Function

' Small helper for the serial path - pokes the status bar so the
' analyst can see which search is currently running.
Private Sub SetSearchStatus(entityLabel As String, searchType As String)
    Application.StatusBar = "OSINT: " & entityLabel & " -> " & searchType & "..."
End Sub

' Flips Excel into its fastest, least-distracting state while the
' macro runs (no calc, no screen redraw, no dialogs) and saves
' the original values so we can restore them on the way out.
Public Sub AutoPrepareRuntime()
    On Error Resume Next
    If m_runtimeApplied Then Exit Sub

    With Application
        m_origCalc = .Calculation
        m_origScreenUpdate = .ScreenUpdating
        m_origEvents = .EnableEvents
        m_origAlerts = .DisplayAlerts
        m_origCursor = .Cursor
        m_origPrintComm = .PrintCommunication

        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Cursor = xlWait
        .PrintCommunication = False
    End With

    m_runtimeApplied = True
    On Error GoTo 0

    TuneBackgroundApps
End Sub

Public Sub AutoRestoreRuntime()
    On Error Resume Next

    ' Always try to restore priorities, even if Excel state wasn't
    ' applied - a half-way failure still might have tuned a few PIDs.
    RestoreBackgroundApps

    If Not m_runtimeApplied Then
        Application.StatusBar = False
        Exit Sub
    End If

    With Application
        .Calculation = m_origCalc
        .ScreenUpdating = m_origScreenUpdate
        .EnableEvents = m_origEvents
        .DisplayAlerts = m_origAlerts
        .Cursor = m_origCursor
        .PrintCommunication = m_origPrintComm
        .StatusBar = False
    End With

    m_runtimeApplied = False
    On Error GoTo 0
End Sub

' If a previous OSINT run died before AutoRestoreRuntime fired
' (VBE Reset, "End" on a runtime error, Excel killed mid-batch...)
' Excel stays in our muted mode: Calculation = manual, EnableEvents
' off, DisplayAlerts off, etc. In Excel 365 that shows up on every
' formula as a yellow "Stale" badge with the value struck through,
' which is alarming on hidden backend sheets the analyst can't fix
' from the ribbon.
'
' This sub force-clears that residue. It only intervenes when it
' sees clear OSINT-shaped fingerprints (our status-bar prefix, or
' the trifecta of muted Excel toggles) so it never accidentally
' overrides a user who genuinely keeps Excel on manual calc.
Private Sub ForceClearStaleOSINTState()
    On Error Resume Next

    Dim looksStale As Boolean
    looksStale = False

    ' Smoking gun #1: our own status-bar message is still up.
    Dim sb As Variant
    sb = Application.StatusBar
    If VarType(sb) = vbString Then
        If InStr(1, CStr(sb), "OSINT", vbTextCompare) > 0 Then looksStale = True
    End If

    ' Smoking gun #2: the trifecta of toggles we mute together.
    ' A normal user would never have all three flipped at once.
    If Not looksStale Then
        If Application.Calculation = xlCalculationManual And _
           Application.EnableEvents = False And _
           Application.DisplayAlerts = False Then
            looksStale = True
        End If
    End If

    If Not looksStale Then Exit Sub

    LogStep "RECOVER: stale OSINT state detected on entry - force-clearing " & _
            "(calc/events/alerts/cursor/screen/print/statusbar)"

    With Application
        .Calculation = xlAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Cursor = xlDefault
        .PrintCommunication = True
        .StatusBar = False
    End With

    ' Lift any background-app priorities the dead run lowered.
    ' Safe no-op if nothing is currently tuned down.
    RestoreBackgroundApps

    ' Reset our flag so the new run's AutoPrepareRuntime starts
    ' from a clean baseline (it'll save the freshly-restored state
    ' as the "original" to put back later).
    m_runtimeApplied = False

    On Error GoTo 0
End Sub

' Public, manual-recovery entry point. Bind it to a button or run
' it from Alt+F8 if the workbook ever ends up with "Stale" cell
' indicators / events not firing / formulas not recalculating
' after an OSINT run that didn't tear down cleanly.
'
' Safe to run any time - it's a no-op if there's nothing to clean.
Public Sub OSINT_RecoverState()
    On Error Resume Next

    Dim took As String

    ' Authoritative path: we have saved originals from a still-live
    ' AutoPrepareRuntime, restore them exactly.
    If m_runtimeApplied Then
        AutoRestoreRuntime
        took = "Restored Excel state from saved originals."
    Else
        ' Fallback path: no saved originals, but state may still
        ' look stuck. ForceClearStaleOSINTState only fires if it
        ' sees real OSINT residue.
        ForceClearStaleOSINTState
        took = "Force-cleared any leftover OSINT state. " & _
               "(If Excel was already healthy, nothing changed.)"
    End If

    ' Always-safe housekeeping regardless of which path ran.
    Application.StatusBar = False
    RestoreBackgroundApps

    ' Nudge a recalc so any "Stale" badges clear immediately
    ' instead of waiting for the next edit.
    On Error Resume Next
    Application.CalculateFull
    On Error GoTo 0

    MsgBox took & vbCrLf & vbCrLf & _
           "Calculation, events, alerts, screen updates, status bar and " & _
           "background-app priorities are back to defaults. " & _
           "Stale-cell indicators have been cleared.", _
           vbInformation, "OSINT Recover"

    On Error GoTo 0
End Sub

' Drop Teams / Outlook / OneDrive to BelowNormal so they stop
' fighting Edge for the 2 vCPUs. Remembers the PIDs we actually
' managed to change, so RestoreBackgroundApps can put them back.
' Uses WMI, which is usable on locked-down VDIs without admin.
' Silently no-ops if WMI is unavailable or the toggle is off.
Private Sub TuneBackgroundApps()
    If Not TUNE_BACKGROUND_PRIORITIES Then Exit Sub

    Set m_tunedPids = New Collection
    On Error Resume Next

    Dim wmi As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If wmi Is Nothing Then Exit Sub

    ' Process names we want to de-prioritise. Upper/lower case
    ' matches what Task Manager / WMI actually report.
    Dim targets As Variant
    targets = Array( _
        "Teams.exe", "ms-teams.exe", "msteams.exe", _
        "msedgewebview2.exe", _
        "OUTLOOK.EXE", _
        "OneDrive.exe")

    Const PRIORITY_BELOW_NORMAL As Long = 16384
    Dim target As Variant
    Dim procs As Object, p As Object
    Dim lowered As Long

    For Each target In targets
        Set procs = wmi.ExecQuery( _
            "SELECT ProcessId FROM Win32_Process WHERE Name = '" & _
            Replace(CStr(target), "'", "''") & "'")
        If Not procs Is Nothing Then
            For Each p In procs
                If p.SetPriority(PRIORITY_BELOW_NORMAL) = 0 Then
                    m_tunedPids.Add CLng(p.ProcessId)
                    lowered = lowered + 1
                End If
            Next p
        End If
    Next target

    If lowered > 0 Then
        LogStep "BG TUNE: lowered " & lowered & " background process(es) to BelowNormal"
    End If
End Sub

' Lifts every PID we tuned back to Normal. Processes that have
' exited in the meantime are just ignored.
Private Sub RestoreBackgroundApps()
    If m_tunedPids Is Nothing Then Exit Sub
    If m_tunedPids.Count = 0 Then
        Set m_tunedPids = Nothing
        Exit Sub
    End If

    On Error Resume Next

    Dim wmi As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If wmi Is Nothing Then
        Set m_tunedPids = Nothing
        Exit Sub
    End If

    Const PRIORITY_NORMAL As Long = 32
    Dim pid As Variant
    Dim procs As Object, p As Object
    Dim restored As Long

    For Each pid In m_tunedPids
        Set procs = wmi.ExecQuery( _
            "SELECT ProcessId FROM Win32_Process WHERE ProcessId = " & CLng(pid))
        If Not procs Is Nothing Then
            For Each p In procs
                If p.SetPriority(PRIORITY_NORMAL) = 0 Then
                    restored = restored + 1
                End If
            Next p
        End If
    Next pid

    LogStep "BG TUNE: restored " & restored & " background process(es) to Normal"
    Set m_tunedPids = Nothing
End Sub

' Fires a throwaway HEAD request to Google so the DNS + TLS
' handshake is already cached by the time Edge gets here. We fire
' and forget - there's no point waiting for the response.
Public Sub AutoPrewarmDNS()
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.SetTimeouts 1500, 2000, 3000, 3000
    http.Open "HEAD", "https://www.google.com/", True   ' True = async
    http.Send
    ' Intentionally NOT calling WaitForResponse - fire and forget.
    Set http = Nothing
    On Error GoTo 0
End Sub

' Pre-warms each worker's headless Edge profile by firing a tiny
' --print-to-pdf of about:blank against every worker's user-data-dir.
' That forces Edge to populate the first-run profile state (Local
' State, Preferences, code caches, GPU cache) so the first real
' search doesn't pay the cold-start tax. Each instance writes a
' throwaway PDF and exits on its own.
'
' Fire-and-forget: we don't wait here. The user prompts that follow
' (negative-news InputBoxes, sigma address dialog, etc.) give the
' prewarm processes plenty of wall-clock time to finish on their own.
' DrainPrewarmAndCleanup is called just before the real dispatch
' starts to make sure no straggler is still holding a profile lock.
Public Sub PrewarmEdgeProfiles()
    On Error Resume Next

    If Not USE_HEADLESS Then Exit Sub

    Dim edgeExe As String
    edgeExe = GetEdgePath
    If Len(edgeExe) = 0 Then Exit Sub
    If Dir(edgeExe) = "" Then Exit Sub

    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim w As Long, profilePath As String, prewarmPdf As String, cmd As String
    Dim fired As Long: fired = 0

    For w = 0 To MAX_PARALLEL - 1
        profilePath = GetWorkerProfilePath(w)
        If Not FSO.FolderExists(profilePath) Then FSO.CreateFolder profilePath

        prewarmPdf = Environ("TEMP") & "\osint_prewarm_w" & w & ".pdf"
        If Dir(prewarmPdf) <> "" Then Kill prewarmPdf

        ' Stripped-down version of the real headless cmd: same flags
        ' that touch the profile dir, none of the search-specific bits.
        cmd = """" & edgeExe & """" & _
              " --headless=new" & _
              " --disable-gpu" & _
              " --user-data-dir=""" & profilePath & """" & _
              " --no-first-run" & _
              " --no-default-browser-check" & _
              " --disable-logging --log-level=3" & _
              " --disable-background-networking" & _
              " --disable-sync" & _
              " --disable-default-apps" & _
              " --disable-component-update" & _
              " --disable-client-side-phishing-detection" & _
              " --metrics-recording-only" & _
              " --mute-audio" & _
              " --print-to-pdf=""" & prewarmPdf & """" & _
              " about:blank"

        wsh.Run cmd, 0, False
        fired = fired + 1
    Next w

    If fired > 0 Then
        LogStep "PREWARM: fired " & fired & " Edge profile warmer(s) - " & _
                "running in background while user is prompted"
    End If

    On Error GoTo 0
End Sub

' Called just before the real dispatch starts. In the common case
' the prewarm Edge processes have already exited on their own (the
' prewarm PDFs exist on disk), so this is a no-op cleanup. If the
' user blew through prompts faster than expected we wait a short
' grace period and then kill any straggler holding a profile lock.
' Even a partially-warmed profile still gets us most of the
' cold-start saving, since the profile dir was populated.
Private Sub DrainPrewarmAndCleanup()
    On Error Resume Next

    Dim w As Long, prewarmPdf As String
    Dim drainStartT As Single: drainStartT = Timer
    Dim allDone As Boolean

    Do
        allDone = True
        For w = 0 To MAX_PARALLEL - 1
            prewarmPdf = Environ("TEMP") & "\osint_prewarm_w" & w & ".pdf"
            If Dir(prewarmPdf) = "" Then
                allDone = False
                Exit For
            End If
        Next w
        If allDone Then Exit Do
        ' 2 s ceiling is plenty - print-to-pdf of about:blank is sub-second
        ' on every VDI we've seen. Beyond that we just kill and move on.
        If Timer - drainStartT > 2 Then Exit Do
        SmartWait 0.1, True
    Loop

    ' Belt-and-braces: kill any straggler msedge.exe still bound to
    ' our worker profiles. No-op if prewarm already exited cleanly.
    For w = 0 To MAX_PARALLEL - 1
        KillEdgeWorkerProcesses w
    Next w

    ' Drop the throwaway prewarm PDFs - they were just the trigger
    ' that got Edge to print-and-exit, the real searches don't need them.
    For w = 0 To MAX_PARALLEL - 1
        prewarmPdf = Environ("TEMP") & "\osint_prewarm_w" & w & ".pdf"
        If Dir(prewarmPdf) <> "" Then Kill prewarmPdf
    Next w

    LogStep "PREWARM: drain done in " & Format(Timer - drainStartT, "0.00") & "s" & _
            " (allDone=" & allDone & ")"

    On Error GoTo 0
End Sub

' Nudge Excel to the top so the next InputBox / MsgBox isn't hidden
' behind something. Cheap insurance against the classic VDI bug
' where a prompt shows up off-screen and the macro looks frozen.
Private Sub ForcePromptToFront()
    On Error Resume Next
    AppActivate Application.Caption
    DoEvents
    On Error GoTo 0
End Sub

' Direct Win32 MessageBoxW call with TOPMOST + SETFOREGROUND.
' VBA's own MsgBox (even with vbSystemModal) loses the z-order
' fight on Citrix / V2 Cloud, so we go around it and talk to
' user32 directly. Returns Win32 IDOK / IDYES / IDNO.
Private Function TopMostMsgBox(ByVal msg As String, _
                               ByVal title As String, _
                               ByVal flags As Long) As Long
    Beep   ' audible cue for the analyst
    On Error Resume Next
    AppActivate Application.Caption
    DoEvents
    On Error GoTo 0
    TopMostMsgBox = MessageBoxW(0, StrPtr(msg), StrPtr(title), _
                                flags Or MB_TOPMOST Or MB_SETFOREGROUND Or MB_SYSTEMMODAL)
End Function

' Returns the window-class of the current foreground window.
' Edge/Chromium windows are "Chrome_WidgetWin_1"; Excel's main
' window is "XLMAIN"; the Windows file dialog is "#32770". We use
' this to make sure focus-dependent keystrokes can only ever reach
' the intended window, never the workbook or the web page.
Private Function ForegroundClassName() As String
    On Error Resume Next
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    Dim buf As String, n As Long
    h = GetForegroundWindow()
    buf = String$(256, vbNullChar)
    n = GetClassName(h, buf, 255)
    ForegroundClassName = Left$(buf, n)
    On Error GoTo 0
End Function

'------------------------------------------------------------------
' File-based step log. Writes one timestamped line per call to
' %TEMP%\osint_debug.log. Survives VBA crashes, hangs, Excel kills.
' Open the file from "Run -> %TEMP%" to diagnose.
'------------------------------------------------------------------
Public Sub LogStep(ByVal msg As String)
    On Error Resume Next
    Dim fnum As Integer, logPath As String, attempt As Long
    logPath = Environ("TEMP") & "\osint_debug.log"
    ' Retry up to 3x with Shared lock so Notepad/Notepad++ holding the
    ' file open cannot silently swallow log lines.
    For attempt = 1 To 3
        fnum = FreeFile
        Err.Clear
        Open logPath For Append Shared As #fnum
        If Err.Number = 0 Then
            Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
            Close #fnum
            Exit For
        End If
    Next attempt
    On Error GoTo 0
End Sub

Public Function BuildEdgeArgs(ByVal url As String) As String
    ' Only the interactive fallback (USE_HEADLESS = False) uses this.
    ' Flags picked to play nicely with V2 Cloud's no-GPU environment.
    BuildEdgeArgs = _
        "--disable-gpu --disable-gpu-compositing --disable-software-rasterizer " & _
        "--disable-background-timer-throttling --disable-renderer-backgrounding " & _
        "--disable-backgrounding-occluded-windows " & _
        "--disable-features=CalculateNativeWinOcclusion,VizDisplayCompositor " & _
        "--renderer-process-limit=2 " & _
        "--no-first-run --no-default-browser-check --disable-sync --disable-extensions " & _
        "--lang=en-US """ & url & """"
End Function

' Reorders the task queue so the statistically heavier searches
' run first. Based on past run logs the fat-tail outliers are
' almost always "Negative News" queries (20-70 s each, vs 10-15 s
' for a plain name search), so if they land at the end of the
' queue one of the workers is stuck on a 60 s task while the other
' sits idle. Launching them first means by the time we get to the
' tail both workers are on short tasks and finish together.
'
' Stable sort - within a weight bucket, original order is kept so
' per-entity ordering is mostly preserved.
Private Function SortTasksHeaviestFirst(ByVal tasks As Collection) As Collection
    Dim n As Long: n = tasks.Count
    If n <= 1 Then
        Set SortTasksHeaviestFirst = tasks
        Exit Function
    End If

    Dim weights() As Long
    Dim ordered() As Variant
    ReDim weights(1 To n)
    ReDim ordered(1 To n)

    Dim i As Long, a As Variant, desc As String
    For i = 1 To n
        a = tasks(i)
        desc = LCase$(CStr(a(3)))
        Select Case True
            Case InStr(desc, "negative news") > 0: weights(i) = 10
            Case InStr(desc, "name + address") > 0: weights(i) = 5
            Case Else:                              weights(i) = 1
        End Select
        ordered(i) = a
    Next i

    ' Insertion sort - descending, stable. n is small (under ~30).
    Dim j As Long, wKey As Long, aKey As Variant
    For i = 2 To n
        wKey = weights(i)
        aKey = ordered(i)
        j = i - 1
        Do While j >= 1
            If weights(j) < wKey Then
                weights(j + 1) = weights(j)
                ordered(j + 1) = ordered(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        weights(j + 1) = wKey
        ordered(j + 1) = aKey
    Next i

    Dim result As New Collection
    For i = 1 To n
        result.Add ordered(i)
    Next i

    LogStep "SORT: reordered " & n & " tasks (heavy first) - heaviest weight=" & _
            weights(1) & ", lightest=" & weights(n)

    Set SortTasksHeaviestFirst = result
End Function

' ============================ Search engine ============================
'
' Runs a flat collection of search tasks. Each task is
' Array(query, savePath, pageNum, description, entityLabel).
'
' In headless mode we keep up to MAX_PARALLEL Edge processes busy at
' once across the whole batch - there's no per-entity boundary, so
' workers are never idle while some other entity's tasks are still
' queued. Anything that comes back suspiciously small (probably a
' CAPTCHA) or that times out gets one automatic retry, requeued
' into the same pool so it overlaps the tail of the initial dispatch
' instead of running as a sequential second pass.
'
' In interactive mode we just run them one at a time.
'
' Returns the number of tasks that produced a valid PDF on disk.
Private Function RunSearchBatch(tasks As Collection, wsh As Object) As Long
    If tasks Is Nothing Then Exit Function
    If tasks.Count = 0 Then Exit Function

    If USE_HEADLESS Then
        RunSearchBatch = RunSearchBatchParallel(tasks, wsh)
    Else
        RunSearchBatch = RunSearchBatchSerial(tasks, wsh)
    End If
End Function

Private Function RunSearchBatchParallel(tasks As Collection, wsh As Object) As Long
    On Error Resume Next

    Dim n As Long: n = tasks.Count
    If n = 0 Then Exit Function

    Dim batchStartT As Single: batchStartT = Timer
    LogStep "BATCH ALL start, n=" & n & ", maxParallel=" & MAX_PARALLEL

    ' Make sure the profile pre-warm Edge processes have either
    ' exited cleanly or been killed before we try to grab the
    ' worker user-data-dirs. Usually they're long gone by now -
    ' this is a sub-50 ms no-op in the typical case.
    DrainPrewarmAndCleanup

    ' Single main pass with inline retries (max 1 per task). Tasks
    ' that come back as CAPTCHA-size or that time out are requeued
    ' into the same pool so the retry happens on the next free
    ' worker - no separate pass, no 5 s cooldown, the work overlaps
    ' the tail of the initial dispatch instead of being sequential.
    Dim mainSizes() As Long
    DispatchTaskPass tasks, wsh, "main", mainSizes, False, TASK_TIMEOUT_SEC, 1

    ' Rescue check: did anything still end up missing or sub-MIN?
    ' Run those one more time at the longer retry timeout. This is
    ' rare in practice (the inline retry already absorbed most of it)
    ' but it's our final correctness net.
    Dim rescueList As Collection
    Set rescueList = New Collection
    Dim i As Long, a As Variant
    For i = 1 To n
        a = tasks(i)
        On Error Resume Next
        If Dir(CStr(a(1))) = "" Or FileLen(CStr(a(1))) <= MIN_PDF_SIZE_BYTES Then
            rescueList.Add tasks(i)
        End If
        On Error GoTo 0
    Next i

    If rescueList.Count > 0 And Not bAbort Then
        LogStep "RESCUE PASS: " & rescueList.Count & " files still missing - final retry"
        Application.StatusBar = "OSINT: final rescue pass for " & rescueList.Count & " missing file(s)..."
        Dim rescueSizes() As Long
        DispatchTaskPass rescueList, wsh, "rescue", rescueSizes, True, RETRY_TIMEOUT_SEC, 0
    End If

    ' Final tally is taken straight from disk - that way any retry
    ' getting the file to a good state still gets counted.
    Dim finalSuccess As Long
    For i = 1 To n
        a = tasks(i)
        On Error Resume Next
        If FileLen(CStr(a(1))) > MIN_PDF_SIZE_BYTES Then
            finalSuccess = finalSuccess + 1
        End If
        On Error GoTo 0
    Next i

    LogStep "BATCH ALL end, success=" & finalSuccess & "/" & n & _
            " wall=" & CLng(Timer - batchStartT) & "s rescue=" & rescueList.Count

    RunSearchBatchParallel = finalSuccess
End Function

' Runs one pass over a task list using the parallel worker pool.
' Fills finalSizes(1..n) with the size each task's PDF ends up at
' on disk, so the caller can decide if a retry is warranted.
'
' When preserveExisting is True (used on retry / rescue passes) we
' don't delete the existing PDF before relaunching Edge - instead
' we rename it to a .retrybak sidecar. If the retry produces a
' good file we drop the backup; if it fails we put the old file
' back. That way we never regress a file that was already working.
Private Sub DispatchTaskPass(tasks As Collection, _
                             wsh As Object, _
                             ByVal passLabel As String, _
                             ByRef finalSizes() As Long, _
                             Optional ByVal preserveExisting As Boolean = False, _
                             Optional ByVal timeoutSeconds As Long = 0, _
                             Optional ByVal maxRetries As Long = 0)
    On Error GoTo Fail

    Dim n As Long: n = tasks.Count
    If n = 0 Then
        ReDim finalSizes(1 To 1)   ' guard against uninitialised array in caller
        Exit Sub
    End If

    Dim qry() As String, pth() As String, pg() As Long
    Dim dsc() As String, lbl() As String
    If timeoutSeconds <= 0 Then timeoutSeconds = TASK_TIMEOUT_SEC

    ' Per-task bookkeeping:
    '   st         - 0 pending / 1 running / 2 done / 3 failed
    '   launchT    - when we kicked Edge off
    '   firstSeenT - when the PDF first appeared on disk with any bytes
    '   wk         - which worker slot the task is using
    '   bak        - path of the preserved old PDF, if any
    '   lastSize / stableHits - watch the file size settle before we
    '                           call the task done, so we never catch
    '                           Edge mid-write and mark a partial PDF.
    '   durSecArr  - captured duration per task for the end-of-pass stats.
    Dim st() As Long
    Dim launchT() As Date
    Dim firstSeenT() As Date
    Dim wk() As Long
    Dim bak() As String
    Dim hadBackup() As Boolean
    Dim lastSize() As Long
    Dim stableHits() As Long
    Dim durSecArr() As Long
    Dim retriesUsed() As Long      ' how many retry attempts already burned per task
    ReDim qry(1 To n), pth(1 To n), pg(1 To n), dsc(1 To n), lbl(1 To n)
    ReDim st(1 To n), launchT(1 To n), firstSeenT(1 To n), wk(1 To n)
    ReDim bak(1 To n), hadBackup(1 To n), lastSize(1 To n), stableHits(1 To n)
    ReDim durSecArr(1 To n), retriesUsed(1 To n)
    ReDim finalSizes(1 To n)

    ' Track how much of the pass each worker was actually busy -
    ' lets us spot a lone worker holding up the whole batch at the tail.
    Dim workerBusyT() As Date
    Dim workerBusySec() As Long
    ReDim workerBusyT(0 To MAX_PARALLEL - 1)
    ReDim workerBusySec(0 To MAX_PARALLEL - 1)
    Dim queueStartT As Date: queueStartT = Now

    Dim i As Long, a As Variant
    For i = 1 To n
        a = tasks(i)
        qry(i) = CStr(a(0))
        pth(i) = CStr(a(1))
        pg(i) = CLng(a(2))
        dsc(i) = CStr(a(3))
        If UBound(a) >= 4 Then lbl(i) = CStr(a(4)) Else lbl(i) = ""
    Next i

    Dim busy() As Boolean
    ReDim busy(0 To MAX_PARALLEL - 1)

    Dim doneCount As Long, runningCount As Long, successCount As Long
    Dim w As Long, pendingFound As Boolean
    Dim cmd As String
    Dim sz As Long
    Dim passStartT As Single: passStartT = Timer
    Dim taskDurSec As Long
    Dim currLabel As String

    ' Totals for "how long did the dispatcher sit idle" - lets us see
    ' whether the sleep/stagger knobs are well-tuned.
    Dim launchCount As Long: launchCount = 0
    Dim retryCount As Long: retryCount = 0    ' inline-retry firings inside this pass
    Dim somethingCompletedThisCycle As Boolean
    Dim idleSleepMs As Long: idleSleepMs = 0
    Dim taskTimeoutSec As Long

    LogStep "PASS [" & passLabel & "] start, n=" & n & ", maxRetries=" & maxRetries

    Do
        somethingCompletedThisCycle = False

        ' Fill any free worker slots with pending tasks. The launch
        ' stagger only fires for the initial ramp-up: once every
        ' worker has been launched once we stop staggering, so a
        ' mid-batch completion gets an instant relaunch.
        For w = 0 To MAX_PARALLEL - 1
            If bAbort Then Exit For
            If Not busy(w) Then
                pendingFound = False
                For i = 1 To n
                    If bAbort Then Exit For
                    If st(i) = 0 Then
                        On Error Resume Next
                        If preserveExisting Then
                            bak(i) = pth(i) & ".retrybak"
                            If Dir(bak(i)) <> "" Then Kill bak(i)
                            If Dir(pth(i)) <> "" Then
                                Name pth(i) As bak(i)
                                hadBackup(i) = True
                            End If
                        Else
                            If Dir(pth(i)) <> "" Then Kill pth(i)
                        End If
                        On Error GoTo Fail

                        cmd = BuildHeadlessEdgeCmd(qry(i), pth(i), pg(i), w)
                        wsh.Run cmd, 0, False

                        st(i) = 1
                        launchT(i) = Now
                        workerBusyT(w) = Now         ' worker utilisation clock starts
                        wk(i) = w
                        busy(w) = True
                        runningCount = runningCount + 1
                        pendingFound = True
                        launchCount = launchCount + 1

                        LogStep "  TASK launch [" & i & "/" & n & "] w" & w & _
                                " running=" & runningCount & _
                                " [" & lbl(i) & "] " & dsc(i)

                        ' Stagger only during initial fill, not mid-batch.
                        If launchCount <= MAX_PARALLEL Then
                            SmartWait t(LAUNCH_STAGGER_SEC), True
                        End If
                        Exit For
                    End If
                Next i
                If Not pendingFound Then Exit For
            End If
        Next w

        ' Poll for completions or timeouts
        For i = 1 To n
            If st(i) = 1 Then
                sz = 0
                If Dir(pth(i)) <> "" Then
                    On Error Resume Next
                    sz = FileLen(pth(i))
                    On Error GoTo Fail
                End If

                ' Remember the moment the PDF first shows up with any
                ' bytes - that's our proxy for Edge having finished
                ' booting and navigated to the SERP.
                If sz > 0 And firstSeenT(i) = #12:00:00 AM# Then
                    firstSeenT(i) = Now
                End If

                If sz > MIN_PDF_SIZE_BYTES Then
                    If sz = lastSize(i) Then
                        stableHits(i) = stableHits(i) + 1
                    Else
                        lastSize(i) = sz
                        stableHits(i) = 0
                    End If

                    If stableHits(i) < 1 Then GoTo ContinuePollingTask

                    ' Tiny PDF on first attempt = almost certainly a
                    ' CAPTCHA / interstitial. If we still have a retry
                    ' available, requeue this task into the same pool
                    ' (no separate retry pass, no cooldown). On the
                    ' last allowed attempt we accept whatever we got.
                    If sz < CAPTCHA_SIZE_HINT And retriesUsed(i) < maxRetries Then
                        On Error Resume Next
                        If Dir(pth(i)) <> "" Then Kill pth(i)
                        On Error GoTo Fail

                        workerBusySec(wk(i)) = workerBusySec(wk(i)) + _
                            DateDiff("s", workerBusyT(wk(i)), Now)
                        st(i) = 0                       ' back to pending
                        retriesUsed(i) = retriesUsed(i) + 1
                        runningCount = runningCount - 1
                        busy(wk(i)) = False
                        somethingCompletedThisCycle = True
                        retryCount = retryCount + 1
                        firstSeenT(i) = #12:00:00 AM#   ' reset bookkeeping
                        lastSize(i) = 0
                        stableHits(i) = 0
                        LogStep "  TASK retry [" & i & "/" & n & "] CAPTCHA-size " & sz & _
                                "B - requeued (attempt " & (retriesUsed(i) + 1) & ")"
                        GoTo ContinuePollingTask
                    End If

                    st(i) = 2
                    finalSizes(i) = sz
                    successCount = successCount + 1
                    doneCount = doneCount + 1
                    runningCount = runningCount - 1
                    busy(wk(i)) = False
                    somethingCompletedThisCycle = True
                    taskDurSec = DateDiff("s", launchT(i), Now)
                    durSecArr(i) = taskDurSec
                    workerBusySec(wk(i)) = workerBusySec(wk(i)) + _
                        DateDiff("s", workerBusyT(wk(i)), Now)
                    currLabel = lbl(i)

                    ' Slice total duration into wait / Edge-start / render
                    ' so the log tells us which phase ate the time.
                    Dim edgeStartSec As Long, renderSec As Long, waitSec As Long
                    waitSec = DateDiff("s", queueStartT, launchT(i))
                    If firstSeenT(i) <> #12:00:00 AM# Then
                        edgeStartSec = DateDiff("s", launchT(i), firstSeenT(i))
                        renderSec = DateDiff("s", firstSeenT(i), Now)
                    Else
                        edgeStartSec = -1
                        renderSec = -1
                    End If

                    LogStep "  TASK done  [" & i & "/" & n & "] w" & wk(i) & _
                            " wait=" & waitSec & "s edgeStart=" & edgeStartSec & _
                            "s render=" & renderSec & "s dur=" & taskDurSec & _
                            "s size=" & sz & "B running=" & runningCount

                    If hadBackup(i) Then
                        On Error Resume Next
                        If Dir(bak(i)) <> "" Then Kill bak(i)
                        On Error GoTo Fail
                    End If

                    If sz < CAPTCHA_SIZE_HINT Then
                        LogStep "  TASK warn  [" & i & "/" & n & "] small PDF " & sz & _
                                "B (possible CAPTCHA, no retries left) - [" & lbl(i) & "] " & dsc(i)
                    End If

                Else
                    ' Adaptive timeout: a task that's already been retried
                    ' once gets the longer RETRY_TIMEOUT_SEC because
                    ' those tend to be the genuinely-slow Google queries.
                    If retriesUsed(i) > 0 Then
                        taskTimeoutSec = RETRY_TIMEOUT_SEC
                    Else
                        taskTimeoutSec = timeoutSeconds
                    End If

                    If DateDiff("s", launchT(i), Now) > taskTimeoutSec Then
                        ' Always free the worker first - the orphan
                        ' Edge process needs killing whether we retry
                        ' or give up.
                        KillEdgeWorkerProcesses wk(i)
                        workerBusySec(wk(i)) = workerBusySec(wk(i)) + _
                            DateDiff("s", workerBusyT(wk(i)), Now)
                        durSecArr(i) = DateDiff("s", launchT(i), Now)

                        Dim toFileState As String
                        If firstSeenT(i) = #12:00:00 AM# Then
                            toFileState = "NO FILE EVER"
                        Else
                            toFileState = "file seen after " & _
                                DateDiff("s", launchT(i), firstSeenT(i)) & "s"
                        End If

                        If retriesUsed(i) < maxRetries Then
                            ' Inline retry: drop straight back into pending,
                            ' the next free worker will pick it up.
                            On Error Resume Next
                            If Not preserveExisting Then
                                If Dir(pth(i)) <> "" Then Kill pth(i)
                            End If
                            On Error GoTo Fail

                            st(i) = 0
                            retriesUsed(i) = retriesUsed(i) + 1
                            runningCount = runningCount - 1
                            busy(wk(i)) = False
                            somethingCompletedThisCycle = True
                            retryCount = retryCount + 1
                            firstSeenT(i) = #12:00:00 AM#
                            lastSize(i) = 0
                            stableHits(i) = 0
                            LogStep "  TASK retry [" & i & "/" & n & "] TIMEOUT after " & _
                                    taskTimeoutSec & "s (" & toFileState & _
                                    ") - requeued (attempt " & (retriesUsed(i) + 1) & ")"
                        Else
                            ' Out of retries - mark as final failure.
                            st(i) = 3
                            finalSizes(i) = 0
                            doneCount = doneCount + 1
                            runningCount = runningCount - 1
                            busy(wk(i)) = False
                            somethingCompletedThisCycle = True
                            If hadBackup(i) Then
                                On Error Resume Next
                                If Dir(pth(i)) <> "" Then Kill pth(i)
                                If Dir(bak(i)) <> "" Then Name bak(i) As pth(i)
                                If Dir(pth(i)) <> "" Then finalSizes(i) = FileLen(pth(i))
                                On Error GoTo Fail
                            End If
                            LogStep "  TASK TIMEOUT [" & i & "/" & n & "] w" & wk(i) & _
                                    " after " & taskTimeoutSec & "s (" & toFileState & _
                                    ") - [" & lbl(i) & "] " & dsc(i)
                        End If
                    End If
                End If
ContinuePollingTask:
            End If
        Next i

        ' Status bar: show pass + most-recent entity + progress
        If currLabel <> "" Then
            Application.StatusBar = "OSINT [" & passLabel & "]: " & currLabel & " - " & _
                doneCount & "/" & n & " saved, " & runningCount & " running"
        Else
            Application.StatusBar = "OSINT [" & passLabel & "]: " & _
                doneCount & "/" & n & " saved, " & runningCount & " running"
        End If

        If bAbort Then Exit Do

        ' If a task just finished, loop back immediately so the launch
        ' phase can push work onto the now-free worker without waiting
        ' out a poll cycle. Only sleep when nothing changed this round.
        If Not somethingCompletedThisCycle Then
            SmartWait t(POLL_INTERVAL_SEC)
            idleSleepMs = idleSleepMs + CLng(t(POLL_INTERVAL_SEC) * 1000)
        End If
    Loop While doneCount < n

    ' End-of-pass stats: how were the task durations distributed,
    ' and how saturated was each worker?
    Dim passWallSec As Long: passWallSec = CLng(Timer - passStartT)
    Dim minD As Long, maxD As Long, sumD As Long, countD As Long
    Dim b0_10 As Long, b10_20 As Long, b20_30 As Long, b30_45 As Long, b45_60 As Long, b60p As Long
    minD = 99999
    For i = 1 To n
        If durSecArr(i) > 0 Then
            countD = countD + 1
            sumD = sumD + durSecArr(i)
            If durSecArr(i) < minD Then minD = durSecArr(i)
            If durSecArr(i) > maxD Then maxD = durSecArr(i)
            Select Case durSecArr(i)
                Case Is < 10:  b0_10 = b0_10 + 1
                Case Is < 20:  b10_20 = b10_20 + 1
                Case Is < 30:  b20_30 = b20_30 + 1
                Case Is < 45:  b30_45 = b30_45 + 1
                Case Is < 60:  b45_60 = b45_60 + 1
                Case Else:     b60p = b60p + 1
            End Select
        End If
    Next i
    If countD = 0 Then minD = 0

    LogStep "PASS [" & passLabel & "] end, success=" & successCount & "/" & n & _
            " wall=" & passWallSec & "s"
    If countD > 0 Then
        LogStep "  DIST [" & passLabel & "] min=" & minD & "s avg=" & (sumD \ countD) & _
                "s
