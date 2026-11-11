Attribute VB_Name = "modWeeklyReport"
'==============================================================================
' Echo Corp - Weekly operations report
'
' One click (GenerateWeeklyReport):
'   1. Reads three raw exports: Raw_Leads, Raw_Invoices, Raw_Web
'   2. Cleans them: mixed date formats, "R 4 500,00" text amounts, messy
'      casing and spacing, duplicate rows, blank or invalid rows
'   3. Works out this week's KPIs against last week
'   4. Rebuilds the Report sheet: KPI table, breakdowns, 8-week trend + chart
'   5. Saves the report as a PDF next to the workbook
'
' Replaces a manual process that took about 3 hours every week.
' Author: Riyasat Rehman - github.com/RiyasatR23
'==============================================================================
Option Explicit

Private Const SH_CONTROL As String = "Control"
Private Const SH_LEADS As String = "Raw_Leads"
Private Const SH_INV As String = "Raw_Invoices"
Private Const SH_WEB As String = "Raw_Web"
Private Const SH_REPORT As String = "Report"
Private Const TREND_WEEKS As Long = 8

' ---- cleaned data, held in memory (arrays are much faster than cell-by-cell) ----
Private nLeads As Long
Private LDate() As Date, LSource() As String, LService() As String
Private LScore() As String, LStatus() As String, LValue() As Double

Private nInv As Long
Private IDate() As Date, IService() As String, IAmount() As Double, IPaid() As Boolean

Private nWeb As Long
Private WDate() As Date, WSessions() As Double, WEnquiries() As Double, WResp() As Double

' ---- data-quality counters shown at the bottom of the report ----
Private rawLeads As Long, dupLeads As Long, badLeads As Long
Private rawInv As Long, dupInv As Long, badInv As Long
Private rawWeb As Long, badWeb As Long
Private pdfProblem As String

Private Type WeekStats
    Leads As Long
    Hot As Long
    Won As Long
    Pipeline As Double      ' estimated value of new leads
    Invoiced As Double
    PaidSoFar As Double     ' part of that week's invoices already paid
    Sessions As Double
    Enquiries As Double
    RespSum As Double
    RespDays As Long
End Type

'==============================================================================
'  ENTRY POINT  (assigned to the "Generate weekly report" button)
'==============================================================================
Public Sub GenerateWeeklyReport()
    Dim started As Double, weekStart As Date, pdfPath As String
    started = Timer

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Cleaning raw data..."

    CheckHeaders SH_LEADS, Array("created_at", "name", "email", "source", "service", "score", "status", "est_value")
    CheckHeaders SH_INV, Array("invoice_no", "invoice_date", "client", "service", "amount", "status")
    CheckHeaders SH_WEB, Array("date", "sessions", "enquiries", "avg_response_hrs")

    LoadLeads
    LoadInvoices
    LoadWeb
    If nLeads + nInv + nWeb = 0 Then
        Err.Raise vbObjectError + 513, , "There is no data in the raw sheets yet. Paste the exports in and try again."
    End If

    weekStart = PickWeek()
    Application.StatusBar = "Building report for week of " & Format$(weekStart, "d mmm yyyy") & "..."
    BuildReport weekStart, started

    Application.StatusBar = "Saving PDF..."
    pdfPath = SavePdf(weekStart)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ThisWorkbook.Worksheets(SH_REPORT).Range("A1").Select
    MsgBox "Weekly report for " & WeekLabel(weekStart) & " is ready" & _
           " (" & Format$(Timer - started, "0.0") & " s)." & vbCrLf & vbCrLf & _
           IIf(Len(pdfPath) > 0, "PDF saved to:" & vbCrLf & pdfPath, "PDF not saved: " & pdfProblem), _
           vbInformation, "Echo weekly report"
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "The report could not be built." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Echo weekly report"
End Sub

'==============================================================================
'  LOAD + CLEAN
'==============================================================================
Private Sub LoadLeads()
    Dim v As Variant, i As Long, d As Variant, email As String, key As String
    Dim seen As Collection
    Set seen = New Collection
    nLeads = 0: rawLeads = 0: dupLeads = 0: badLeads = 0

    v = ReadTable(SH_LEADS, 8)
    If IsEmpty(v) Then Exit Sub
    ReDim LDate(1 To UBound(v, 1)): ReDim LSource(1 To UBound(v, 1)): ReDim LService(1 To UBound(v, 1))
    ReDim LScore(1 To UBound(v, 1)): ReDim LStatus(1 To UBound(v, 1)): ReDim LValue(1 To UBound(v, 1))

    For i = 1 To UBound(v, 1)
        If RowIsBlank(v, i) Then GoTo NextRow
        rawLeads = rawLeads + 1

        d = ParseDate(v(i, 1))
        email = LCase$(CleanText(v(i, 3)))
        If IsEmpty(d) Or InStr(email, "@") = 0 Then
            badLeads = badLeads + 1
            GoTo NextRow
        End If

        ' the same person submitting the form twice on one day counts once
        key = email & "|" & Format$(d, "yyyymmdd")
        If InCollection(seen, key) Then
            dupLeads = dupLeads + 1
            GoTo NextRow
        End If
        seen.Add True, key

        nLeads = nLeads + 1
        LDate(nLeads) = d
        LSource(nLeads) = NormSource(v(i, 4))
        LService(nLeads) = IIf(Len(CleanText(v(i, 5))) = 0, "Other", CleanText(v(i, 5)))
        LScore(nLeads) = NormScore(v(i, 6))
        LStatus(nLeads) = LCase$(CleanText(v(i, 7)))
        LValue(nLeads) = ParseMoney(v(i, 8))
NextRow:
    Next i
End Sub

Private Sub LoadInvoices()
    Dim v As Variant, i As Long, d As Variant, invNo As String
    Dim seen As Collection
    Set seen = New Collection
    nInv = 0: rawInv = 0: dupInv = 0: badInv = 0

    v = ReadTable(SH_INV, 6)
    If IsEmpty(v) Then Exit Sub
    ReDim IDate(1 To UBound(v, 1)): ReDim IService(1 To UBound(v, 1))
    ReDim IAmount(1 To UBound(v, 1)): ReDim IPaid(1 To UBound(v, 1))

    For i = 1 To UBound(v, 1)
        If RowIsBlank(v, i) Then GoTo NextRow
        rawInv = rawInv + 1

        invNo = UCase$(CleanText(v(i, 1)))
        d = ParseDate(v(i, 2))
        If IsEmpty(d) Or Len(invNo) = 0 Then
            badInv = badInv + 1
            GoTo NextRow
        End If
        If InCollection(seen, invNo) Then          ' invoice exported twice
            dupInv = dupInv + 1
            GoTo NextRow
        End If
        seen.Add True, invNo

        nInv = nInv + 1
        IDate(nInv) = d
        IService(nInv) = IIf(Len(CleanText(v(i, 4))) = 0, "Other", CleanText(v(i, 4)))
        IAmount(nInv) = ParseMoney(v(i, 5))
        IPaid(nInv) = (LCase$(CleanText(v(i, 6))) = "paid")
NextRow:
    Next i
End Sub

Private Sub LoadWeb()
    Dim v As Variant, i As Long, d As Variant
    nWeb = 0: rawWeb = 0: badWeb = 0

    v = ReadTable(SH_WEB, 4)
    If IsEmpty(v) Then Exit Sub
    ReDim WDate(1 To UBound(v, 1)): ReDim WSessions(1 To UBound(v, 1))
    ReDim WEnquiries(1 To UBound(v, 1)): ReDim WResp(1 To UBound(v, 1))

    For i = 1 To UBound(v, 1)
        If RowIsBlank(v, i) Then GoTo NextRow
        rawWeb = rawWeb + 1
        d = ParseDate(v(i, 1))
        If IsEmpty(d) Then
            badWeb = badWeb + 1
            GoTo NextRow
        End If
        nWeb = nWeb + 1
        WDate(nWeb) = d
        WSessions(nWeb) = ToNumber(v(i, 2))
        WEnquiries(nWeb) = ToNumber(v(i, 3))
        WResp(nWeb) = ToNumber(v(i, 4))
NextRow:
    Next i
End Sub

'==============================================================================
'  CALCULATIONS
'==============================================================================
Private Function GetStats(ByVal weekStart As Date) As WeekStats
    Dim s As WeekStats, i As Long, weekEnd As Date
    weekEnd = weekStart + 6

    For i = 1 To nLeads
        If LDate(i) >= weekStart And LDate(i) <= weekEnd Then
            s.Leads = s.Leads + 1
            s.Pipeline = s.Pipeline + LValue(i)
            If LScore(i) = "hot" Then s.Hot = s.Hot + 1
            If LStatus(i) = "won" Then s.Won = s.Won + 1
        End If
    Next i
    For i = 1 To nInv
        If IDate(i) >= weekStart And IDate(i) <= weekEnd Then
            s.Invoiced = s.Invoiced + IAmount(i)
            If IPaid(i) Then s.PaidSoFar = s.PaidSoFar + IAmount(i)
        End If
    Next i
    For i = 1 To nWeb
        If WDate(i) >= weekStart And WDate(i) <= weekEnd Then
            s.Sessions = s.Sessions + WSessions(i)
            s.Enquiries = s.Enquiries + WEnquiries(i)
            If WResp(i) > 0 Then
                s.RespSum = s.RespSum + WResp(i)
                s.RespDays = s.RespDays + 1
            End If
        End If
    Next i
    GetStats = s
End Function

' Week to report on: the date in Control!C4, otherwise the latest full Mon-Sun week in the data
Private Function PickWeek() As Date
    Dim d As Variant, latest As Date, ws As Date
    d = ParseDate(ThisWorkbook.Worksheets(SH_CONTROL).Range("C4").Value)
    If Not IsEmpty(d) Then
        PickWeek = WeekStartOf(CDate(d))
        Exit Function
    End If
    latest = LatestDate()
    ws = WeekStartOf(latest)
    If latest - ws < 6 Then ws = ws - 7          ' latest week isn't finished yet
    PickWeek = ws
End Function

Private Function LatestDate() As Date
    Dim i As Long, m As Date
    For i = 1 To nLeads
        If LDate(i) > m Then m = LDate(i)
    Next i
    For i = 1 To nInv
        If IDate(i) > m Then m = IDate(i)
    Next i
    For i = 1 To nWeb
        If WDate(i) > m Then m = WDate(i)
    Next i
    LatestDate = m
End Function

Private Function WeekStartOf(ByVal d As Date) As Date
    WeekStartOf = DateSerial(Year(d), Month(d), Day(d)) - Weekday(d, vbMonday) + 1
End Function

Private Function WeekLabel(ByVal weekStart As Date) As String
    WeekLabel = Format$(weekStart, "ddd d mmm") & " - " & Format$(weekStart + 6, "ddd d mmm yyyy")
End Function

'==============================================================================
'  REPORT
'==============================================================================
Private Sub BuildReport(ByVal weekStart As Date, ByVal started As Double)
    Dim sh As Worksheet, co As ChartObject
    Dim cur As WeekStats, prev As WeekStats
    Dim r As Long, i As Long, k As Long

    Set sh = ThisWorkbook.Worksheets(SH_REPORT)
    For Each co In sh.ChartObjects
        co.Delete
    Next co
    sh.Cells.Clear
    sh.Cells.Font.Name = "Calibri"
    sh.Cells.Font.Size = 10
    sh.Activate
    ActiveWindow.DisplayGridlines = False

    cur = GetStats(weekStart)
    prev = GetStats(weekStart - 7)

    ' ---- header ----
    sh.Range("B2").Value = "Echo Corp " & ChrW(183) & " Weekly operations report"
    sh.Range("B2").Font.Size = 18
    sh.Range("B2").Font.Bold = True
    sh.Range("B2").Font.Color = Brand()
    sh.Range("B3").Value = WeekLabel(weekStart) & "   " & ChrW(183) & "   compared with " & WeekLabel(weekStart - 7)
    sh.Range("B3").Font.Color = RGB(100, 100, 110)

    ' ---- KPI table ----
    r = 5
    HeaderRow sh, r, Array("Metric", "This week", "Last week", "Change", "% change")
    KpiRow sh, r + 1, "New leads", cur.Leads, prev.Leads, "0", True, False
    KpiRow sh, r + 2, "Hot leads", cur.Hot, prev.Hot, "0", True, False
    KpiRow sh, r + 3, "Leads won", cur.Won, prev.Won, "0", True, False
    KpiRow sh, r + 4, "Pipeline value (new leads)", cur.Pipeline, prev.Pipeline, MoneyFmt(), True, False
    KpiRow sh, r + 5, "Revenue invoiced", cur.Invoiced, prev.Invoiced, MoneyFmt(), True, False
    KpiRow sh, r + 6, "  of which paid so far", cur.PaidSoFar, prev.PaidSoFar, MoneyFmt(), True, False
    KpiRow sh, r + 7, "Website sessions", cur.Sessions, prev.Sessions, "#,##0", True, False
    KpiRow sh, r + 8, "Enquiry rate (enquiries / sessions)", SafeDiv(cur.Enquiries, cur.Sessions), _
           SafeDiv(prev.Enquiries, prev.Sessions), "0.0%", True, True
    KpiRow sh, r + 9, "Avg first response (hours)", SafeDiv(cur.RespSum, cur.RespDays), _
           SafeDiv(prev.RespSum, prev.RespDays), "0.0", False, False
    sh.Range(sh.Cells(r + 1, 2), sh.Cells(r + 9, 6)).Borders(xlEdgeBottom).Color = RGB(220, 220, 228)
    r = r + 11

    ' ---- outstanding ----
    sh.Cells(r, 2).Value = "Unpaid invoices up to " & Format$(weekStart + 6, "d mmm")
    sh.Cells(r, 3).Value = Outstanding(weekStart + 6)
    sh.Cells(r, 3).NumberFormat = MoneyFmt()
    sh.Cells(r, 2).Font.Bold = True
    sh.Cells(r, 3).Font.Bold = True
    sh.Cells(r, 3).Font.Color = RGB(185, 28, 28)
    r = r + 2

    ' ---- breakdowns ----
    r = ServiceTable(sh, r, weekStart) + 1
    r = SourceTable(sh, r, weekStart) + 1

    ' ---- 8-week trend ----
    Dim trendTop As Long, s As WeekStats, wk As Date
    SectionTitle sh, r, "Last " & TREND_WEEKS & " weeks"
    HeaderRow sh, r + 1, Array("Week of", "New leads", "Hot leads", "Invoiced", "Sessions")
    trendTop = r + 2
    For k = TREND_WEEKS - 1 To 0 Step -1
        wk = weekStart - 7 * k
        s = GetStats(wk)
        i = trendTop + (TREND_WEEKS - 1 - k)
        sh.Cells(i, 2).Value = "'" & Format$(wk, "d mmm")
        sh.Cells(i, 3).Value = s.Leads
        sh.Cells(i, 4).Value = s.Hot
        sh.Cells(i, 5).Value = s.Invoiced
        sh.Cells(i, 5).NumberFormat = MoneyFmt()
        sh.Cells(i, 6).Value = s.Sessions
        sh.Cells(i, 6).NumberFormat = "#,##0"
    Next k
    sh.Range(sh.Cells(trendTop + TREND_WEEKS - 1, 2), sh.Cells(trendTop + TREND_WEEKS - 1, 6)).Font.Bold = True
    r = trendTop + TREND_WEEKS + 1

    ' ---- data quality ----
    SectionTitle sh, r, "Data cleaned"
    sh.Cells(r + 1, 2).Value = "Leads: " & rawLeads & " rows read, " & Plural(dupLeads, "duplicate") & " removed, " & badLeads & " invalid skipped"
    sh.Cells(r + 2, 2).Value = "Invoices: " & rawInv & " rows read, " & Plural(dupInv, "duplicate") & " removed, " & badInv & " invalid skipped"
    sh.Cells(r + 3, 2).Value = "Website: " & rawWeb & " days read, " & badWeb & " invalid skipped"
    sh.Range(sh.Cells(r + 1, 2), sh.Cells(r + 3, 2)).Font.Color = RGB(100, 100, 110)
    sh.Cells(r + 5, 2).Value = "Generated " & Format$(Now, "d mmm yyyy hh:mm") & " in " & Format$(Timer - started, "0.0") & " s by GenerateWeeklyReport"
    sh.Cells(r + 5, 2).Font.Italic = True
    sh.Cells(r + 5, 2).Font.Color = RGB(150, 150, 160)

    ' ---- layout ----
    sh.Columns("A").ColumnWidth = 2
    sh.Columns("B").ColumnWidth = 36
    sh.Columns("C:F").ColumnWidth = 14
    sh.Columns("G").ColumnWidth = 3
    sh.Range("C:F").HorizontalAlignment = xlRight

    AddTrendChart sh, trendTop, trendTop + TREND_WEEKS - 1
    SetupPage sh, r + 5
End Sub

Private Sub KpiRow(sh As Worksheet, ByVal r As Long, ByVal label As String, ByVal cur As Double, ByVal prev As Double, _
                   ByVal fmt As String, ByVal higherIsBetter As Boolean, ByVal isRate As Boolean)
    Dim diff As Double
    diff = cur - prev
    sh.Cells(r, 2).Value = label
    sh.Cells(r, 3).Value = cur
    sh.Cells(r, 4).Value = prev
    sh.Cells(r, 3).NumberFormat = fmt
    sh.Cells(r, 4).NumberFormat = fmt
    sh.Cells(r, 3).Font.Bold = True
    sh.Cells(r, 4).Font.Color = RGB(110, 110, 120)

    If isRate Then                                ' a change in a rate is shown in percentage points
        sh.Cells(r, 5).Value = diff * 100
        sh.Cells(r, 5).NumberFormat = "+0.0"" pp"";-0.0"" pp"";0.0"" pp"""
    Else
        sh.Cells(r, 5).Value = diff
        sh.Cells(r, 5).NumberFormat = "+" & fmt & ";-" & fmt & ";" & fmt
        If prev <> 0 Then
            sh.Cells(r, 6).Value = diff / Abs(prev)
            sh.Cells(r, 6).NumberFormat = "+0%;-0%;0%"
        ElseIf cur <> 0 Then
            sh.Cells(r, 6).Value = "new"
        End If
    End If

    If Abs(diff) > 0.000001 Then
        If (diff > 0) = higherIsBetter Then
            sh.Range(sh.Cells(r, 5), sh.Cells(r, 6)).Font.Color = RGB(22, 128, 61)
        Else
            sh.Range(sh.Cells(r, 5), sh.Cells(r, 6)).Font.Color = RGB(185, 28, 28)
        End If
    End If
End Sub

Private Function ServiceTable(sh As Worksheet, ByVal r As Long, ByVal weekStart As Date) As Long
    Dim names() As String, n As Long, i As Long, j As Long, row As Long
    Dim leads As Long, hot As Long, pipe As Double, inv As Double
    ReDim names(1 To nLeads + nInv + 1)

    For i = 1 To nLeads
        If LDate(i) >= weekStart And LDate(i) <= weekStart + 6 Then AddUnique names, n, LService(i)
    Next i
    For i = 1 To nInv
        If IDate(i) >= weekStart And IDate(i) <= weekStart + 6 Then AddUnique names, n, IService(i)
    Next i

    SectionTitle sh, r, "By service (this week)"
    HeaderRow sh, r + 1, Array("Service", "New leads", "Hot leads", "Pipeline", "Invoiced")
    row = r + 2
    For j = 1 To n
        leads = 0: hot = 0: pipe = 0: inv = 0
        For i = 1 To nLeads
            If LService(i) = names(j) And LDate(i) >= weekStart And LDate(i) <= weekStart + 6 Then
                leads = leads + 1
                pipe = pipe + LValue(i)
                If LScore(i) = "hot" Then hot = hot + 1
            End If
        Next i
        For i = 1 To nInv
            If IService(i) = names(j) And IDate(i) >= weekStart And IDate(i) <= weekStart + 6 Then inv = inv + IAmount(i)
        Next i
        sh.Cells(row, 2).Value = names(j)
        sh.Cells(row, 3).Value = leads
        sh.Cells(row, 4).Value = hot
        sh.Cells(row, 5).Value = pipe
        sh.Cells(row, 6).Value = inv
        sh.Range(sh.Cells(row, 5), sh.Cells(row, 6)).NumberFormat = MoneyFmt()
        row = row + 1
    Next j
    If n > 1 Then SortBlock sh, r + 2, row - 1, 3
    ServiceTable = row
End Function

Private Function SourceTable(sh As Worksheet, ByVal r As Long, ByVal weekStart As Date) As Long
    Dim names() As String, n As Long, i As Long, j As Long, row As Long, cnt As Long, total As Long, hot As Long
    ReDim names(1 To nLeads + 1)
    For i = 1 To nLeads
        If LDate(i) >= weekStart And LDate(i) <= weekStart + 6 Then
            AddUnique names, n, LSource(i)
            total = total + 1
        End If
    Next i

    SectionTitle sh, r, "Where leads came from (this week)"
    HeaderRow sh, r + 1, Array("Source", "New leads", "Share", "Hot leads", "")
    row = r + 2
    For j = 1 To n
        cnt = 0: hot = 0
        For i = 1 To nLeads
            If LSource(i) = names(j) And LDate(i) >= weekStart And LDate(i) <= weekStart + 6 Then
                cnt = cnt + 1
                If LScore(i) = "hot" Then hot = hot + 1
            End If
        Next i
        sh.Cells(row, 2).Value = names(j)
        sh.Cells(row, 3).Value = cnt
        sh.Cells(row, 4).Value = SafeDiv(cnt, total)
        sh.Cells(row, 4).NumberFormat = "0%"
        sh.Cells(row, 5).Value = hot
        row = row + 1
    Next j
    If n > 1 Then SortBlock sh, r + 2, row - 1, 3
    SourceTable = row
End Function

Private Function Outstanding(ByVal upTo As Date) As Double
    Dim i As Long
    For i = 1 To nInv
        If Not IPaid(i) And IDate(i) <= upTo Then Outstanding = Outstanding + IAmount(i)
    Next i
End Function

Private Sub AddTrendChart(sh As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim co As ChartObject
    Set co = sh.ChartObjects.Add(sh.Range("H5").Left, sh.Range("H5").Top, 470, 280)
    co.Name = "TrendChart"
    With co.Chart
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .ChartType = xlColumnClustered
            .Name = "Invoiced (R)"
            .Values = sh.Range(sh.Cells(firstRow, 5), sh.Cells(lastRow, 5))
            .XValues = sh.Range(sh.Cells(firstRow, 2), sh.Cells(lastRow, 2))
            .Format.Fill.ForeColor.RGB = Brand()
        End With
        With .SeriesCollection.NewSeries
            .Name = "New leads"
            .Values = sh.Range(sh.Cells(firstRow, 3), sh.Cells(lastRow, 3))
            .XValues = sh.Range(sh.Cells(firstRow, 2), sh.Cells(lastRow, 2))
            .ChartType = xlLineMarkers
            .AxisGroup = xlSecondary
            .Format.Line.ForeColor.RGB = RGB(245, 158, 11)
            .MarkerBackgroundColor = RGB(245, 158, 11)
            .MarkerForegroundColor = RGB(245, 158, 11)
        End With
        .HasTitle = True
        .ChartTitle.Text = "Revenue invoiced and new leads, last " & TREND_WEEKS & " weeks"
        .ChartTitle.Font.Size = 12
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "General"
    End With
End Sub

Private Sub SetupPage(sh As Worksheet, ByVal lastRow As Long)
    Dim lastCol As Long
    lastCol = sh.ChartObjects("TrendChart").BottomRightCell.Column
    If sh.ChartObjects("TrendChart").BottomRightCell.Row > lastRow Then lastRow = sh.ChartObjects("TrendChart").BottomRightCell.Row
    On Error Resume Next                      ' page setup fails if no printer driver is installed; not fatal
    With sh.PageSetup
        .PrintArea = sh.Range(sh.Cells(1, 1), sh.Cells(lastRow, lastCol)).Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
    End With
    On Error GoTo 0
End Sub

' Saves the Report sheet as a PDF. Returns the path, or "" if it could not be saved.
Private Function SavePdf(ByVal weekStart As Date) As String
    Dim folder As String, f As String
    folder = ThisWorkbook.Path
    If Len(folder) = 0 Then
        pdfProblem = "save the workbook first, then run it again."
        Exit Function
    End If
    If LCase$(Left$(folder, 4)) = "http" Then folder = Application.DefaultFilePath   ' OneDrive/SharePoint URL
    f = folder & Application.PathSeparator & "Echo weekly report " & Format$(weekStart, "yyyy-mm-dd") & ".pdf"

    On Error GoTo NoPdf
    ThisWorkbook.Worksheets(SH_REPORT).ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    SavePdf = f
    Exit Function
NoPdf:
    SavePdf = ""
    pdfProblem = Err.Description
    Err.Clear
End Function

'==============================================================================
'  FORMATTING HELPERS
'==============================================================================
Private Function Brand() As Long
    Brand = RGB(75, 43, 214)
End Function

Private Function MoneyFmt() As String
    MoneyFmt = """R ""#,##0"
End Function

Private Sub HeaderRow(sh As Worksheet, ByVal r As Long, headers As Variant)
    Dim i As Long
    For i = 0 To UBound(headers)
        sh.Cells(r, 2 + i).Value = headers(i)
    Next i
    With sh.Range(sh.Cells(r, 2), sh.Cells(r, 6))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = Brand()
    End With
End Sub

Private Sub SectionTitle(sh As Worksheet, ByVal r As Long, ByVal title As String)
    sh.Cells(r, 2).Value = title
    sh.Cells(r, 2).Font.Bold = True
    sh.Cells(r, 2).Font.Size = 12
    sh.Cells(r, 2).Font.Color = Brand()
End Sub

Private Sub SortBlock(sh As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal sortCol As Long)
    sh.Range(sh.Cells(firstRow, 2), sh.Cells(lastRow, 6)).Sort _
        Key1:=sh.Cells(firstRow, sortCol), Order1:=xlDescending, Header:=xlNo
End Sub

'==============================================================================
'  CLEANING HELPERS
'==============================================================================
Private Sub CheckHeaders(ByVal sheetName As String, expected As Variant)
    Dim sh As Worksheet, i As Long, got As String
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then Err.Raise vbObjectError + 514, , "The sheet """ & sheetName & """ is missing."
    For i = 0 To UBound(expected)
        got = LCase$(CleanText(sh.Cells(1, i + 1).Value))
        If got <> expected(i) Then
            Err.Raise vbObjectError + 515, , sheetName & ": column " & Chr$(65 + i) & " should be """ & expected(i) & _
                """ but is """ & got & """." & vbCrLf & "Paste the export with its header row into A1."
        End If
    Next i
End Sub

Private Function ReadTable(ByVal sheetName As String, ByVal nCols As Long) As Variant
    Dim sh As Worksheet, lastRow As Long
    Set sh = ThisWorkbook.Worksheets(sheetName)
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ReadTable = Empty
    Else
        ReadTable = sh.Range(sh.Cells(2, 1), sh.Cells(lastRow, nCols)).Value
    End If
End Function

Private Function RowIsBlank(v As Variant, ByVal i As Long) As Boolean
    Dim c As Long
    For c = 1 To UBound(v, 2)
        If Len(CleanText(v(i, c))) > 0 Then Exit Function
    Next c
    RowIsBlank = True
End Function

' Trims normal and non-breaking spaces; error cells become ""
Private Function CleanText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    CleanText = Application.WorksheetFunction.Trim(Replace(CStr(v), Chr$(160), " "))
End Function

' Accepts real Excel dates, ISO text (2026-11-03 or 2026-11-03T09:14:22Z),
' yyyy/mm/dd and South African dd/mm/yyyy. Returns Empty if it isn't a date.
Private Function ParseDate(ByVal v As Variant) As Variant
    Dim s As String, y As Long, m As Long, d As Long
    ParseDate = Empty
    If IsError(v) Or IsEmpty(v) Then Exit Function

    If VarType(v) = vbDate Then
        ParseDate = DateSerial(Year(v), Month(v), Day(v))
        Exit Function
    End If
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        If v > 30000 And v < 80000 Then ParseDate = CDate(Int(v))   ' an Excel date serial
        Exit Function
    End If

    s = CleanText(v)
    If Len(s) < 10 Then Exit Function
    If Mid$(s, 5, 1) Like "[-/]" And Mid$(s, 8, 1) Like "[-/]" Then            ' yyyy-mm-dd
        If Not (Left$(s, 4) Like "####" And Mid$(s, 6, 2) Like "##" And Mid$(s, 9, 2) Like "##") Then Exit Function
        y = CLng(Left$(s, 4)): m = CLng(Mid$(s, 6, 2)): d = CLng(Mid$(s, 9, 2))
    ElseIf Mid$(s, 3, 1) Like "[-/.]" And Mid$(s, 6, 1) Like "[-/.]" Then       ' dd/mm/yyyy
        If Not (Left$(s, 2) Like "##" And Mid$(s, 4, 2) Like "##" And Mid$(s, 7, 4) Like "####") Then Exit Function
        d = CLng(Left$(s, 2)): m = CLng(Mid$(s, 4, 2)): y = CLng(Mid$(s, 7, 4))
    Else
        Exit Function
    End If
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    ParseDate = DateSerial(y, m, d)
End Function

' Accepts 4500, "4500", "R4 500", "R 4 500,00", "R4500.00", "R4,500"
Private Function ParseMoney(ByVal v As Variant) As Double
    Dim s As String, p As Long
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then
        If IsNumeric(v) Then ParseMoney = CDbl(v)
        Exit Function
    End If
    s = UCase$(CleanText(v))
    s = Replace(Replace(Replace(s, "R", ""), " ", ""), Chr$(160), "")
    p = InStrRev(s, ",")
    If p > 0 And InStr(s, ".") = 0 And Len(s) - p = 2 Then
        s = Replace(s, ",", ".")                    ' 4500,00  -> decimal comma
    Else
        s = Replace(s, ",", "")                     ' 4,500    -> thousands separator
    End If
    ParseMoney = Val(s)
End Function

Private Function ToNumber(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then ToNumber = CDbl(v) Else ToNumber = Val(CleanText(v))
End Function

Private Function NormSource(ByVal v As Variant) As String
    Select Case LCase$(CleanText(v))
        Case "website", "web", "site", "contact form": NormSource = "Website"
        Case "whatsapp", "wa": NormSource = "WhatsApp"
        Case "referral", "referred", "word of mouth": NormSource = "Referral"
        Case "instagram", "ig", "insta": NormSource = "Instagram"
        Case "google business", "gbp", "google", "google maps": NormSource = "Google Business"
        Case "": NormSource = "Unknown"
        Case Else: NormSource = StrConv(LCase$(CleanText(v)), vbProperCase)
    End Select
End Function

Private Function NormScore(ByVal v As Variant) As String
    Select Case LCase$(CleanText(v))
        Case "hot", "warm", "cold": NormScore = LCase$(CleanText(v))
        Case Else: NormScore = "unscored"
    End Select
End Function

Private Function InCollection(c As Collection, ByVal key As String) As Boolean
    Dim x As Variant
    On Error Resume Next
    x = c.Item(key)
    InCollection = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub AddUnique(arr() As String, n As Long, ByVal value As String)
    Dim i As Long
    For i = 1 To n
        If arr(i) = value Then Exit Sub
    Next i
    n = n + 1
    arr(n) = value
End Sub

Private Function Plural(ByVal n As Long, ByVal word As String) As String
    Plural = n & " " & word & IIf(n = 1, "", "s")
End Function

Private Function SafeDiv(ByVal a As Double, ByVal b As Double) As Double
    If b <> 0 Then SafeDiv = a / b
End Function
