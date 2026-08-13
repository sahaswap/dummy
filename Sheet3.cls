Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim stepName As String
    On Error GoTo ErrorHandler
    
    stepName = "0. Initialization"
    
    ' Break sheet grouping if active (grouped sheets prevent Protect/Unprotect)
    On Error Resume Next
    If Not ActiveWindow Is Nothing Then
        If ActiveWindow.SelectedSheets.count > 1 Then
            Me.Select
        End If
    End If
    On Error GoTo ErrorHandler
    
    LogDebug "Worksheet_Change triggered on sheet '" & Me.Name & "' for cells: " & Target.Address
    LogDebug "Initial State - Sheet Protected: " & Me.ProtectContents & ", Workbook Protected: " & ThisWorkbook.ProtectStructure
    
    ' ==========================================================
    ' 0. THE TOP BREAD: UNLOCK EVERYTHING FIRST
    ' ==========================================================
    stepName = "1. Unprotecting Workbook and Sheet"
    ThisWorkbook.Unprotect Password:="p7ss"
    Me.Unprotect Password:="p7ss"
    LogDebug "After Unprotect - Sheet Protected: " & Me.ProtectContents & ", Workbook Protected: " & ThisWorkbook.ProtectStructure

    ' --- 1. DEFINE YOUR RANGES ---
    stepName = "2. Defining Ranges"
    Dim DecisionCell As Range, TemplateCell As Range
    Dim ProgNameRange As Range, EscScenarioRange As Range
    
    Dim CustNameCell As Range, CustAddressCell As Range
    Dim CpNameRange As Range, CpAddressRange As Range
    Dim AllAutoCapCells As Range
    
    ' Logic Ranges
    Set DecisionCell = Range("J6")
    Set TemplateCell = Range("J7")
    Set ProgNameRange = Range("J8:T8")
    Set EscScenarioRange = Range("J9:T9")
    
    ' Autocapitalize Ranges
    ' (fixed off-by-one: Name value is J14, Address value is J15 - the
    '  old J13/J14 watched the section header + name, so the address at
    '  J15 was never capitalized.)
    Set CustNameCell = Range("J14")
    Set CustAddressCell = Range("J15")
    Set CpNameRange = Range("J19:J24")
    Set CpAddressRange = Range("T19:T24")
    
    ' Bundle all text fields that need capitalizing into one group
    Set AllAutoCapCells = Union(CustNameCell, CustAddressCell, CpNameRange, CpAddressRange)
    
    ' Disable events so our code changes don't trigger an infinite loop
    Application.EnableEvents = False
    
    ' ==========================================================
    ' NEW RULE: SMART AUTOCAPITALIZE (WITH ROMAN NUMERAL EXCEPTIONS)
    ' ==========================================================
    stepName = "3. Checking Autocapitalize"
    ' Check if they typed or pasted into ANY of the name/address cells
    If Not Intersect(Target, AllAutoCapCells) Is Nothing Then
        LogDebug "Autocapitalize triggered"
        Dim changedCell As Range
        Dim words() As String
        Dim i As Integer
        Dim finalName As String
        
        ' Loop through exactly what was changed (handles pasting multiple rows safely)
        For Each changedCell In Intersect(Target, AllAutoCapCells)
            ' Only apply if the cell isn't blank
            If Not IsEmpty(changedCell.Value) Then
                words = Split(changedCell.Value, " ")
                finalName = ""
                For i = LBound(words) To UBound(words)
                    words(i) = Application.WorksheetFunction.Proper(words(i))
                    Select Case UCase(words(i))
                        Case "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "LLC", "INC", "LTD", "LLP", "PLC", "GMBH", "US", "UK"
                            words(i) = UCase(words(i))
                    End Select
                    If finalName = "" Then
                        finalName = words(i)
                    Else
                        finalName = finalName & " " & words(i)
                    End If
                Next i
                changedCell.Value = finalName
            End If
        Next changedCell
    End If

    ' ==========================================================
    ' OLD RULES: GREYING OUT / LOCKING LOGIC
    ' ==========================================================
    stepName = "4. Checking Greying/Locking Logic"
    ' Check if EITHER Decision (J6) or Template (J7) was changed
    If Not Intersect(Target, Union(DecisionCell, TemplateCell)) Is Nothing Then
        Dim decValue As String
        Dim tempValue As String
        
        decValue = UCase(Trim(DecisionCell.Value))
        tempValue = UCase(Trim(TemplateCell.Value))
        LogDebug "Decision/Template changed - Decision: " & decValue & ", Template: " & tempValue

        ' RULE 0 (RFI): Decision = RFI greys + locks + clears BOTH Program Name
        ' (J8:T8) and Escalation Scenario (J9:T9). Skip the Crypto / Esc+Wise rules.
        If decValue = "RFI" Then
            With ProgNameRange.Interior
                .Pattern = xlPatternCrissCross
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.15
            End With
            ProgNameRange.Locked = True
            ProgNameRange.Value = ""
            With EscScenarioRange.Interior
                .Pattern = xlPatternCrissCross
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.15
            End With
            EscScenarioRange.Locked = True
            EscScenarioRange.Value = ""
        Else
            ' RULE 1: Program Name (J8:T8) always unlocked and ungreyed for Escalation/Non-Escalation
            With ProgNameRange.Interior
                .Pattern = xlSolid
                .Color = DecisionCell.Interior.Color
                .TintAndShade = 0
            End With
            ProgNameRange.Locked = False
            
            ' RULE 2: Escalation Scenario (J9:T9) ALWAYS locked since there are no dropdowns
            With EscScenarioRange.Interior
                .Pattern = xlPatternCrissCross
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.15
            End With
            EscScenarioRange.Locked = True
            EscScenarioRange.Value = ""
        End If
    End If

    ' ==============================================================================
    ' PART 3: THE J8 / WISE RULE
    ' ==============================================================================
    stepName = "5. Checking J8/Wise Rule"
    If Not Intersect(Target, Range("J6")) Is Nothing Then
        If Range("J6").Value <> "Wise" Then
            Range("J8").Value = ""
        End If
    End If

    ' --- DEFENSIVE ---
    stepName = "6. Checking Defensive Rules"
    If Not Intersect(Target, DecisionCell) Is Nothing Then
        Dim curT As String, okT As Boolean
        curT = UCase(Trim(TemplateCell.Value))
        okT = True
        If decValue = "RFI" Then
            If curT <> "WISE" And curT <> "AIRWALLEX" And curT <> "COMMON" Then okT = False
        Else
            If curT <> "WISE" And curT <> "FLUZ" And curT <> "CRYPTO" Then okT = False
        End If
        If curT = "" Then okT = True
        If Not okT Then TemplateCell.Value = ""
    End If

    ' ==========================================================
    ' NEW RULE: REFRESH SEARCH MATRIX HYPERLINKS
    ' ==========================================================
    stepName = "6b. Refreshing Search Matrix Hyperlinks"
    If Not Intersect(Target, Range("J10:J24,T19:T24")) Is Nothing Then
        RefreshSearchMatrixHyperlinks
    End If

    ' ==========================================================
    ' NEW RULE: COUNTRY <-> ISO CODE TWO-WAY SYNC (country-risk row 27)
    ' Country Name (G27), ISO Code 2 (L27) and ISO Code 3 (O27) are all
    ' searchable dropdowns. Pick any one and the other two ISO fields
    ' fill from the master table O36:T289 (O=Name, P=ISO2, Q=ISO3). Risk
    ' Score / Risk Classes stay formula-driven off G27, so they follow
    ' automatically. Events are already disabled above, so these writes
    ' don't re-trigger this handler.
    ' ==========================================================
    stepName = "6c. Country/ISO two-way sync"
    Dim ctryCell As Range, iso2Cell As Range, iso3Cell As Range
    Dim nameArr As Range, iso2Arr As Range, iso3Arr As Range
    Set ctryCell = Range("G27")
    Set iso2Cell = Range("L27")
    Set iso3Cell = Range("O27")
    Set nameArr = Range("O36:O289")
    Set iso2Arr = Range("P36:P289")
    Set iso3Arr = Range("Q36:Q289")

    If Not Intersect(Target, ctryCell) Is Nothing Then
        ' Country picked -> fill both ISO codes.
        If Trim(ctryCell.Value) = "" Then
            iso2Cell.Value = "": iso3Cell.Value = ""
        Else
            iso2Cell.Value = LookupOrBlank(CStr(ctryCell.Value), nameArr, iso2Arr)
            iso3Cell.Value = LookupOrBlank(CStr(ctryCell.Value), nameArr, iso3Arr)
        End If
    ElseIf Not Intersect(Target, iso2Cell) Is Nothing Then
        ' ISO2 picked -> set country (risk formulas cascade), fill ISO3.
        Dim nm2 As String
        If Trim(iso2Cell.Value) = "" Then
            ctryCell.Value = "": iso3Cell.Value = ""
        Else
            nm2 = LookupOrBlank(CStr(iso2Cell.Value), iso2Arr, nameArr)
            ctryCell.Value = nm2
            If nm2 = "" Then iso3Cell.Value = "" Else iso3Cell.Value = LookupOrBlank(nm2, nameArr, iso3Arr)
        End If
    ElseIf Not Intersect(Target, iso3Cell) Is Nothing Then
        ' ISO3 picked -> set country, fill ISO2.
        Dim nm3 As String
        If Trim(iso3Cell.Value) = "" Then
            ctryCell.Value = "": iso2Cell.Value = ""
        Else
            nm3 = LookupOrBlank(CStr(iso3Cell.Value), iso3Arr, nameArr)
            ctryCell.Value = nm3
            If nm3 = "" Then iso2Cell.Value = "" Else iso2Cell.Value = LookupOrBlank(nm3, nameArr, iso2Arr)
        End If
    End If

    ' ==========================================================
    ' THE BOTTOM BREAD: RELOCK EVERYTHING AND EXIT SAFELY
    ' ==========================================================
    stepName = "7. Protecting Sheet"
    Application.EnableEvents = True
    Me.Protect Password:="p7ss"
    
    stepName = "8. Protecting Workbook"
    ThisWorkbook.Protect Password:="p7ss", Structure:=True
    
    LogDebug "Worksheet_Change completed successfully"
    Exit Sub

ErrorHandler:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.source
    
    LogDebug "ERROR occurred in step '" & stepName & "' - Err: " & errNum & " - Desc: " & errDesc & " - Source: " & errSrc
    
    ' Clean up events before leaving
    Application.EnableEvents = True
    
    ' Diagnostic protection check
    On Error Resume Next
    LogDebug "Attempting emergency relock..."
    Me.Protect Password:="p7ss"
    If Err.Number <> 0 Then
        LogDebug "Emergency Me.Protect failed: " & Err.Description
    Else
        LogDebug "Emergency Me.Protect succeeded"
    End If
    
    Err.Clear
    ThisWorkbook.Protect Password:="p7ss", Structure:=True
    If Err.Number <> 0 Then
        LogDebug "Emergency ThisWorkbook.Protect failed: " & Err.Description
    Else
        LogDebug "Emergency ThisWorkbook.Protect succeeded"
    End If
    On Error GoTo 0
    
    MsgBox "An error occurred in step '" & stepName & "': " & errDesc, vbCritical, "AML Cockpit Debugger"
End Sub

' Match key in lookupCol and return the same-row value from returnCol,
' or "" if not found. Both ranges are single columns of equal height
' (O/P/Q 36:289). Used by the Country <-> ISO two-way sync above.
Private Function LookupOrBlank(ByVal key As String, ByVal lookupCol As Range, ByVal returnCol As Range) As String
    Dim m As Variant
    m = Application.Match(key, lookupCol, 0)
    If IsError(m) Then
        LookupOrBlank = ""
    Else
        LookupOrBlank = CStr(returnCol.Cells(CLng(m), 1).Value)
    End If
End Function

Private Sub LogDebug(ByVal msg As String)
    On Error Resume Next
    Dim fnum As Integer
    Dim logPath As String
    logPath = Environ("TEMP") & "\sheet3_debug.log"
    fnum = FreeFile
    Open logPath For Append Shared As #fnum
    Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
    Close #fnum
    On Error GoTo 0
End Sub
