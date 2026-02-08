Attribute VB_Name = "GreenStarMacros"
Option Explicit

' ============================================================
' GREEN STAR BUILDINGS v1.1 - INTERACTIVE WORKBOOK MACROS
' ============================================================

' ── Data ──
Private Const RULES_DATA As String = "Industry Development|ID.6|ID.5|Yes
Responsible Construction|RC.3|RC.1|Yes
Responsible Construction|RC.2|RC.1|No
Responsible Construction|RC.5|RC.4|Yes
Verification and Handover|VH.6|VH.5|Yes
Verification and Handover|VH.8|VH.7|Yes
Verification and Handover|VH.27|VH.26|Yes
Responsible Resource Mgmt|RRM.3|RRM.2|Yes
Responsible Resource Mgmt|RRM.6|RRM.5|Yes
Responsible Resource Mgmt|RRM.8|RRM.7|Yes
Responsible Resource Mgmt|RRM.13|RRM.12|Yes
Responsible Procurement|RP.13|RP.12|Yes
Impacts Disclosure|ID2.6|ID2.5|Yes
Clean Air|CA.10|CA.9|Yes
Light Quality|LQ.8|LQ.7|Yes
Light Quality|LQ.9|LQ.7|No
Light Quality|LQ.10|LQ.7|No
Light Quality|LQ.11|LQ.7|No
Exposure to Toxins|ET.2|ET.1|Yes
Exposure to Toxins|ET.3|ET.1|Yes
Amenity and Comfort|AmC.4|AmC.3|Yes
Connection to Nature|CN.5|CN.4|Yes
Connection to Nature|CN.6|CN.4|Yes
Climate Resilience|CR.2|CR.1|Yes
Climate Resilience|CR.3|CR.1|Yes
Climate Resilience|CR.4|CR.1|Yes
Operations Resilience|OR.5|OR.4|Yes
Operations Resilience|OR.7|OR.6|Yes
Community Resilience|CoR.2|CoR.1|Yes
Community Resilience|CoR.3|CoR.1|Yes
Community Resilience|CoR.4|CoR.1|Yes
Grid Resilience|GR.2|GR.1|Yes
Grid Resilience|GR.3|GR.1|Yes
Grid Resilience|GR.5|GR.4|Yes
Grid Resilience|GR.6|GR.4|Yes
Grid Resilience|GR.8|GR.7|Yes
Energy Source|ES.6|ES.5|Yes
Upfront Carbon Reduction|UCR.3|UCR.2|Yes
Upfront Carbon Reduction|UCR.4|UCR.2|Yes
Upfront Carbon Reduction|UCR.5|UCR.2|Yes
Water Use|WU.4|WU.3|Yes
Water Use|WU.6|WU.5|Yes
Contribution to Place|CP.2|CP.1|Yes
Culture Heritage Identity|CHI.2|CHI.1|Yes
First Nations Inclusion|FNI.2|FNI.1|Yes
First Nations Inclusion|FNI.3|FNI.1|Yes
Design for Equity|DE.5|DE.4|Yes
Impacts to Nature|IN.2|IN.1|Yes
Impacts to Nature|IN.3|IN.1|Yes
Nature Connectivity|NC.2|NC.1|Yes
Nature Connectivity|NC.6|NC.5|Yes
Nature Stewardship|NS.2|NS.1|Yes
Nature Stewardship|NS.3|NS.1|Yes
Waterway Protection|WP.6|WP.5|Yes
Market Transformation|MT.5|MT.4|Yes"
Private Const META_DATA As String = "Industry Development|Responsible|1F4E28|15
Responsible Construction|Responsible|1F4E28|25
Verification and Handover|Responsible|1F4E28|35
Responsible Resource Mgmt|Responsible|1F4E28|16
Responsible Procurement|Responsible|1F4E28|16
Responsible Structure|Responsible|1F4E28|7
Responsible Envelope|Responsible|1F4E28|5
Responsible Systems|Responsible|1F4E28|5
Responsible Finishes|Responsible|1F4E28|5
Impacts Disclosure|Responsible|1F4E28|8
Clean Air|Healthy|1565C0|12
Light Quality|Healthy|1565C0|12
Acoustic Comfort|Healthy|1565C0|13
Exposure to Toxins|Healthy|1565C0|9
Amenity and Comfort|Healthy|1565C0|7
Connection to Nature|Healthy|1565C0|8
Climate Resilience|Resilient|E65100|8
Operations Resilience|Resilient|E65100|7
Community Resilience|Resilient|E65100|4
Heat Resilience|Resilient|E65100|6
Grid Resilience|Resilient|E65100|8
Energy Source|Positive|2E7D32|10
Energy Use|Positive|2E7D32|9
Upfront Carbon Reduction|Positive|2E7D32|9
Upfront Carbon Compensation|Positive|2E7D32|5
Refrigerant Systems Impacts|Positive|2E7D32|6
Low-Emissions Transport|Positive|2E7D32|5
Design for Circularity|Positive|2E7D32|7
Water Use|Positive|2E7D32|8
Movement and Place|Places|6A1B9A|6
Enjoyable Places|Places|6A1B9A|5
Contribution to Place|Places|6A1B9A|5
Culture Heritage Identity|Places|6A1B9A|4
Inclusive Construction|People|C62828|7
First Nations Inclusion|People|C62828|6
Procurement Workforce Inclusion|People|C62828|5
Design for Equity|People|C62828|8
Impacts to Nature|Nature|00695C|7
Biodiversity Enhancement|Nature|00695C|5
Nature Connectivity|Nature|00695C|6
Nature Stewardship|Nature|00695C|5
Waterway Protection|Nature|00695C|7
Market Transformation|Leadership|F57F17|5
Leadership Challenges|Leadership|F57F17|4"

' ── Colour Scheme ──
Private Const CLR_DARK_BG As Long = &H21201A
Private Const CLR_DARK_CARD As Long = &H342C28
Private Const CLR_DARK_TEXT As Long = &HE0E0E0
Private Const CLR_DARK_INPUT As Long = &H3A3130

Private gDarkMode As Boolean
Private gSearchSheet As String
Private gSearchRow As Long

' ============================================================
' INITIALISATION
' ============================================================
Public Sub InitWorkbook()
    SetupDashboard
End Sub

' ============================================================
' HANDLE CELL CHANGES (Response column = G)
' ============================================================
Public Sub HandleChange(Sh As Object, Target As Range)
    ' Apply conditional visibility rules
    ApplyConditionalRules Sh
    ' Update progress on Dashboard
    UpdateDashboardProgress
    ' Log to history
    LogChange Sh.Name, Target.Row, Target.Value
End Sub

' ============================================================
' CONDITIONAL VISIBILITY
' ============================================================
Public Sub ApplyConditionalRules(Sh As Object)
    Dim rules() As String
    Dim parts() As String
    Dim i As Long
    Dim sName As String

    If Len(RULES_DATA) = 0 Then Exit Sub
    rules = Split(RULES_DATA, vbLf)
    sName = Sh.Name

    For i = LBound(rules) To UBound(rules)
        If Len(rules(i)) = 0 Then GoTo NextRule
        parts = Split(rules(i), "|")
        If UBound(parts) < 3 Then GoTo NextRule
        If parts(0) <> sName Then GoTo NextRule

        Dim followerRef As String, gatewayRef As String, showWhen As String
        followerRef = parts(1)
        gatewayRef = parts(2)
        showWhen = parts(3)

        ' Find gateway row
        Dim gwRow As Long, fRow As Long
        gwRow = FindRefRow(Sh, gatewayRef)
        fRow = FindRefRow(Sh, followerRef)
        If gwRow = 0 Or fRow = 0 Then GoTo NextRule

        Dim gwVal As String
        gwVal = CStr(Sh.Cells(gwRow, 7).Value)

        If gwVal = showWhen Then
            If Sh.Rows(fRow).Hidden Then
                Sh.Rows(fRow).Hidden = False
            End If
        Else
            If Not Sh.Rows(fRow).Hidden Then
                Sh.Rows(fRow).Hidden = True
            End If
        End If
NextRule:
    Next i
End Sub

Private Function FindRefRow(Sh As Object, ref As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = Sh.Cells(Sh.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(Sh.Cells(r, 1).Value) = ref Then
            FindRefRow = r
            Exit Function
        End If
    Next r
    FindRefRow = 0
End Function

Public Sub ApplyAllConditionalRules()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard" And ws.Name <> "History" And ws.Name <> "SearchResults" Then
            ApplyConditionalRules ws
        End If
    Next ws
End Sub

' ============================================================
' DASHBOARD PROGRESS
' ============================================================
Public Sub SetupDashboard()
    On Error Resume Next
    Dim dsh As Worksheet
    Set dsh = ThisWorkbook.Worksheets("Dashboard")
    If dsh Is Nothing Then Exit Sub

    UpdateDashboardProgress
    ApplyAllConditionalRules
    On Error GoTo 0
End Sub

Public Sub UpdateDashboardProgress()
    Dim dsh As Worksheet
    Set dsh = ThisWorkbook.Worksheets("Dashboard")
    If dsh Is Nothing Then Exit Sub

    Dim meta() As String
    Dim parts() As String
    Dim totalQ As Long, totalA As Long
    totalQ = 0: totalA = 0

    If Len(META_DATA) = 0 Then Exit Sub
    meta = Split(META_DATA, vbLf)

    Dim dashRow As Long
    dashRow = 5  ' First credit row on dashboard

    Dim i As Long
    For i = LBound(meta) To UBound(meta)
        If Len(meta(i)) = 0 Then GoTo NextMeta
        parts = Split(meta(i), "|")
        If UBound(parts) < 3 Then GoTo NextMeta

        Dim sName As String, qCount As Long
        sName = parts(0)
        qCount = CLng(parts(3))

        ' Count answered in this credit sheet
        Dim ws As Worksheet, answered As Long, visible As Long
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sName)
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextMeta

        answered = 0: visible = 0
        Dim r As Long
        For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If Len(CStr(ws.Cells(r, 5).Value)) > 0 Then  ' Has question type = is a question row
                If Not ws.Rows(r).Hidden Then
                    visible = visible + 1
                    If Len(CStr(ws.Cells(r, 7).Value)) > 0 Then
                        answered = answered + 1
                    End If
                End If
            End If
        Next r

        totalQ = totalQ + visible
        totalA = totalA + answered

        ' Update dashboard row
        If dashRow <= dsh.Cells(dsh.Rows.Count, 1).End(xlUp).Row + 5 Then
            dsh.Cells(dashRow, 4).Value = answered
            dsh.Cells(dashRow, 5).Value = visible
            If visible > 0 Then
                dsh.Cells(dashRow, 6).Value = answered / visible
            Else
                dsh.Cells(dashRow, 6).Value = 0
            End If
            dashRow = dashRow + 1
        End If
NextMeta:
    Next i

    ' Update totals
    dsh.Cells(2, 4).Value = totalA
    dsh.Cells(2, 5).Value = totalQ
    If totalQ > 0 Then
        dsh.Cells(2, 6).Value = totalA / totalQ
    Else
        dsh.Cells(2, 6).Value = 0
    End If
End Sub

' ============================================================
' N/A TOGGLE
' ============================================================
Public Sub ToggleNA()
    Dim sName As String
    sName = ActiveSheet.Name
    If sName = "Dashboard" Or sName = "History" Or sName = "SearchResults" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Check current state - look at row 2 font color
    If ws.Cells(2, 1).Font.Color = RGB(180, 180, 180) Then
        ' Currently N/A - re-enable
        Dim r2 As Long
        For r2 = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ws.Rows(r2).Font.Color = RGB(0, 0, 0)
        Next r2
        ws.Cells(1, 8).Value = ""
    Else
        ' Mark as N/A
        Dim r3 As Long
        For r3 = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ws.Rows(r3).Font.Color = RGB(180, 180, 180)
        Next r3
        ws.Cells(1, 8).Value = "N/A"
    End If
    UpdateDashboardProgress
End Sub

' ============================================================
' REVIEW MODE - Highlight unanswered
' ============================================================
Public Sub ReviewMode()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.Name = "Dashboard" Or ws.Name = "History" Or ws.Name = "SearchResults" Then Exit Sub

    Dim r As Long, unanswered As Long
    unanswered = 0

    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Len(CStr(ws.Cells(r, 5).Value)) > 0 Then  ' Question row
            If Not ws.Rows(r).Hidden Then
                If Len(CStr(ws.Cells(r, 7).Value)) = 0 Then
                    ' Highlight unanswered
                    ws.Cells(r, 7).Interior.Color = RGB(255, 243, 224)
                    ws.Cells(r, 7).Borders.Color = RGB(255, 152, 0)
                    unanswered = unanswered + 1
                Else
                    ' Clear highlight
                    ws.Cells(r, 7).Interior.Color = RGB(255, 255, 255)
                    ws.Cells(r, 7).Borders.Color = RGB(200, 200, 200)
                End If
            End If
        End If
    Next r

    MsgBox unanswered & " unanswered question(s) highlighted in orange on " & ws.Name, vbInformation, "Review Mode"
End Sub

' ============================================================
' SEARCH
' ============================================================
Public Sub SearchQuestions()
    Dim query As String
    query = InputBox("Search across all questions:" & vbCrLf & vbCrLf & "Enter search term(s):", "Search Green Star Questions")
    If Len(query) = 0 Then Exit Sub

    Dim searchTerm As String
    searchTerm = LCase(Trim(query))

    ' Create or clear SearchResults sheet
    Dim sr As Worksheet
    On Error Resume Next
    Set sr = ThisWorkbook.Worksheets("SearchResults")
    On Error GoTo 0
    If sr Is Nothing Then
        Set sr = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sr.Name = "SearchResults"
    End If
    sr.Cells.Clear

    ' Header
    sr.Cells(1, 1).Value = "Search Results for: """ & query & """"
    sr.Cells(1, 1).Font.Bold = True
    sr.Cells(1, 1).Font.Size = 14

    sr.Cells(2, 1).Value = "Credit"
    sr.Cells(2, 2).Value = "Ref"
    sr.Cells(2, 3).Value = "Question"
    sr.Cells(2, 4).Value = "Type"
    sr.Cells(2, 5).Value = "Current Response"
    Dim c As Long
    For c = 1 To 5
        sr.Cells(2, c).Font.Bold = True
        sr.Cells(2, c).Interior.Color = RGB(31, 78, 40)
        sr.Cells(2, c).Font.Color = RGB(255, 255, 255)
    Next c

    sr.Columns(1).ColumnWidth = 25
    sr.Columns(2).ColumnWidth = 8
    sr.Columns(3).ColumnWidth = 60
    sr.Columns(4).ColumnWidth = 16
    sr.Columns(5).ColumnWidth = 40

    Dim resultRow As Long
    resultRow = 3

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Dashboard" Or ws.Name = "History" Or ws.Name = "SearchResults" Then GoTo NextSheet
        Dim r As Long
        For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If Len(CStr(ws.Cells(r, 5).Value)) > 0 Then
                Dim haystack As String
                haystack = LCase(CStr(ws.Cells(r, 1).Value) & " " & CStr(ws.Cells(r, 6).Value) & " " & CStr(ws.Cells(r, 8).Value))
                If InStr(haystack, searchTerm) > 0 Then
                    sr.Cells(resultRow, 1).Value = ws.Name
                    sr.Cells(resultRow, 2).Value = ws.Cells(r, 1).Value
                    sr.Cells(resultRow, 3).Value = ws.Cells(r, 6).Value
                    sr.Cells(resultRow, 4).Value = ws.Cells(r, 5).Value
                    sr.Cells(resultRow, 5).Value = ws.Cells(r, 7).Value
                    ' Add hyperlink to jump to the question
                    sr.Hyperlinks.Add sr.Cells(resultRow, 2), "", "'" & ws.Name & "'!A" & r, "Go to question"
                    resultRow = resultRow + 1
                End If
            End If
        Next r
NextSheet:
    Next ws

    sr.Cells(1, 3).Value = (resultRow - 3) & " result(s) found"

    sr.Activate
End Sub

' ============================================================
' VERSION HISTORY
' ============================================================
Public Sub LogChange(sheetName As String, row As Long, newValue As Variant)
    On Error Resume Next
    Dim hsh As Worksheet
    Set hsh = ThisWorkbook.Worksheets("History")
    If hsh Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = hsh.Cells(hsh.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 3 Then nextRow = 3

    ' Keep max 500 entries
    If nextRow > 502 Then
        hsh.Rows("3:103").Delete
        nextRow = hsh.Cells(hsh.Rows.Count, 1).End(xlUp).Row + 1
    End If

    hsh.Cells(nextRow, 1).Value = Now
    hsh.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    hsh.Cells(nextRow, 2).Value = sheetName
    hsh.Cells(nextRow, 3).Value = "Row " & row
    hsh.Cells(nextRow, 4).Value = CStr(newValue)
    On Error GoTo 0
End Sub

Public Sub ShowHistory()
    On Error Resume Next
    ThisWorkbook.Worksheets("History").Activate
    On Error GoTo 0
End Sub

' ============================================================
' DARK MODE
' ============================================================
Public Sub ToggleDarkMode()
    gDarkMode = Not gDarkMode
    Dim ws As Worksheet

    If gDarkMode Then
        For Each ws In ThisWorkbook.Worksheets
            ws.Tab.Color = RGB(30, 33, 39)
            Dim lastR As Long, lastC As Long
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastC = 8
            Dim r As Long, cl As Long
            For r = 1 To lastR
                For cl = 1 To lastC
                    If ws.Cells(r, cl).Interior.Color = RGB(255, 255, 255) Or _
                       ws.Cells(r, cl).Interior.ColorIndex = xlNone Then
                        ws.Cells(r, cl).Interior.Color = CLR_DARK_CARD
                        ws.Cells(r, cl).Font.Color = CLR_DARK_TEXT
                    End If
                Next cl
            Next r
        Next ws
        Application.StatusBar = "Dark mode ON"
    Else
        For Each ws In ThisWorkbook.Worksheets
            ws.Tab.Color = xlNone
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastC = 8
            For r = 1 To lastR
                For cl = 1 To lastC
                    If ws.Cells(r, cl).Interior.Color = CLR_DARK_CARD Then
                        ws.Cells(r, cl).Interior.Color = RGB(255, 255, 255)
                        ws.Cells(r, cl).Font.Color = RGB(0, 0, 0)
                    End If
                Next cl
            Next r
        Next ws
        Application.StatusBar = "Dark mode OFF"
    End If
End Sub

' ============================================================
' NAVIGATION HELPERS
' ============================================================
Public Sub GoToDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub

Public Sub RefreshAll()
    ApplyAllConditionalRules
    UpdateDashboardProgress
    MsgBox "All conditional rules applied and dashboard updated.", vbInformation, "Refresh Complete"
End Sub
