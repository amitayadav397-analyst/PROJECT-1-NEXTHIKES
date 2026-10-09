Attribute VB_Name = "BikeSharingAutomation"
'==============================================================================
' Bike Sharing Demand Analysis - Automation module (Project tasks 6 and 7)
' Workbook: nexthikes_project_excel (saved as .xlsm)
'
' WHAT THIS MODULE DOES
'   RunFullReport      One-click macro. Recalculates the workbook, reads the
'                      data table "tblFinal" and rebuilds the sheet
'                      "Automated Report" (KPIs, two tables, two charts,
'                      top-5 busiest hours, data checks, insight sentences
'                      and the list of anomalies).
'   ExportReportPDF    Saves the "Automated Report" sheet as a PDF file in
'                      the same folder as the workbook.
'   HighlightAnomalies Rebuilds only the coloured anomaly list on the report.
'   ResetFilters       Clears table filters and sets the Dashboard filters
'                      back to "All" and the full date range.
'   DayType, WeatherLabel, DemandLevel
'                      Custom functions (UDFs). They can also be used in a
'                      cell, for example =DayType(H2) or =WeatherLabel(I2).
'
' IMPORTANT DESIGN RULE: "read-only on data"
'   Every column in Final_Dataset is a FORMULA linked to Dataset_A and
'   Dataset_3. If a macro sorted, de-duplicated or typed over that table,
'   the formulas would break. So these macros only READ the data and write
'   their results to two new sheets: "Automated Report" and "Automation Log".
'   The cleaning itself was done with formulas (see the Data_Quality sheet).
'
' WORKS ON: Excel 2016 or newer, on Mac and on Windows. It uses only plain
'   VBA (arrays, Collections, loops). No ActiveX, no Windows-only objects.
'==============================================================================
Option Explicit

' ---- Names used in the workbook (change here if you rename something) ----
Private Const DATA_SHEET As String = "Final_Dataset"      ' sheet holding the data table
Private Const DATA_TABLE As String = "tblFinal"           ' Excel Table name
Private Const DASH_SHEET As String = "Dashboard"          ' sheet with the drop-down filters
Private Const ANOM_SHEET As String = "Anomalies"          ' sheet with the Z-score analysis
Private Const ANOM_THRESHOLD_CELL As String = "B4"        ' z threshold typed by the user
Private Const ANOM_FLAG_RANGE As String = "M31:M1030"     ' "High"/"Low" flags on Anomalies
Private Const REPORT_SHEET As String = "Automated Report" ' sheet this module builds
Private Const LOG_SHEET As String = "Automation Log"      ' sheet with the run history

' ---- Where each block starts on the report sheet (row numbers) ----
Private Const ROW_KPI As Long = 5          ' Key numbers
Private Const ROW_CHECKS As Long = 17      ' Data checks
Private Const ROW_HOURLY As Long = 24      ' Hourly summary table (24 rows)
Private Const ROW_WEEKDAY As Long = 52     ' Weekday summary table (7 rows)
Private Const ROW_TOP5 As Long = 63        ' Top 5 busiest hours
Private Const ROW_INSIGHTS As Long = 72    ' Auto-written insight sentences
Private Const ROW_ANOMALY As Long = 85     ' Anomaly list (written by HighlightAnomalies)

' ---- Data loaded from the table (shared by the procedures below) ----
' mData is a 2-D array: mData(row, column). Reading the whole table into an
' array once is much faster than reading 1,000 cells one by one.
Private mData As Variant
Private mRows As Long

' Column numbers inside tblFinal (found by header name, so the order of
' columns can change without breaking the code).
Private colInstant As Long
Private colDteday As Long
Private colHr As Long
Private colWeekday As Long
Private colWeather As Long
Private colCasual As Long
Private colRegistered As Long
Private colCnt As Long

' The 16 original (base) columns that must never be blank.
Private Const BASE_COLUMNS As String = _
    "instant,dteday,season,yr,mnth,hr,holiday,weekday,weathersit,temp,atemp,hum,windspeed,casual,registered,cnt"


'==============================================================================
' 1. MAIN BUTTON MACRO
'==============================================================================
Public Sub RunFullReport()
    Dim oldCalc As Long          ' remembers the user's calculation setting
    Dim errMsg As String         ' filled only if something goes wrong
    Dim startTime As Double

    startTime = Timer
    oldCalc = Application.Calculation
    On Error GoTo Fail

    ' Step 1: refresh any data connections (safe even when there are none).
    RefreshConnectionsSafely

    ' Step 2: recalculate every formula so the table holds current values.
    Application.CalculateFull

    ' Step 3: switch off screen redraw and automatic calculation while we
    '         write the report. This makes the macro faster and stops flicker.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Step 4: read the table into memory, then build the report.
    LoadTableData
    BuildReport
    WriteAnomalyList GetOrCreateSheet(REPORT_SHEET)

CleanExit:
    ' This part ALWAYS runs, even after an error, so Excel is put back to
    ' normal (screen updating on, calculation as it was before).
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True

    If errMsg = "" Then
        WriteLog "RunFullReport", "OK - report rebuilt in " & Format(Timer - startTime, "0.0") & " s"
        ThisWorkbook.Worksheets(REPORT_SHEET).Activate
        ThisWorkbook.Worksheets(REPORT_SHEET).Range("A1").Select
        MsgBox "The Automated Report has been updated.", vbInformation, "Bike Sharing Automation"
    Else
        WriteLog "RunFullReport", "ERROR - " & errMsg
        MsgBox "The report could not be built." & vbNewLine & vbNewLine & _
               "Reason: " & errMsg, vbExclamation, "Bike Sharing Automation"
    End If
    Exit Sub

Fail:
    ' Save the error text, then jump to CleanExit to restore Excel.
    errMsg = Err.Description
    Resume CleanExit
End Sub


'==============================================================================
' 2. EXPORT THE REPORT TO PDF
'==============================================================================
Public Sub ExportReportPDF()
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim errMsg As String

    On Error GoTo Fail

    ' The workbook must be saved first, otherwise it has no folder.
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first (File > Save As > .xlsm), " & _
               "then run ExportReportPDF again.", vbExclamation, "Export to PDF"
        Exit Sub
    End If

    ' The report must exist. If not, build it first.
    If Not SheetExists(REPORT_SHEET) Then RunFullReport
    If Not SheetExists(REPORT_SHEET) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)

    ' Page setup: landscape, fit to one page wide. Wrapped in its own
    ' error trap because page setup needs a printer driver on some Macs.
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo Fail

    ' File name example: Automated_Report_2026-10-09_1430.pdf
    ' Application.PathSeparator is "/" on Mac and "\" on Windows.
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & _
              "Automated_Report_" & Format(Now, "yyyy-mm-dd_hhnn") & ".pdf"

    ' ExportAsFixedFormat with xlTypePDF is the built-in "save as PDF".
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath

    WriteLog "ExportReportPDF", "OK - saved " & pdfPath
    MsgBox "PDF saved here:" & vbNewLine & pdfPath, vbInformation, "Export to PDF"
    Exit Sub

Fail:
    errMsg = Err.Description
    WriteLog "ExportReportPDF", "ERROR - " & errMsg
    MsgBox "The PDF could not be saved." & vbNewLine & vbNewLine & _
           "Reason: " & errMsg & vbNewLine & vbNewLine & _
           "On a Mac, Excel may need permission to use this folder. " & _
           "Click 'Grant Access' if asked, or use File > Save As > PDF instead.", _
           vbExclamation, "Export to PDF"
End Sub


'==============================================================================
' 3. HIGHLIGHT ANOMALIES (writes a coloured list on the report)
'    We do NOT colour the Anomalies sheet itself: it already has conditional
'    formatting (orange rows) and overwriting it could break that.
'==============================================================================
Public Sub HighlightAnomalies()
    Dim oldCalc As Long
    Dim errMsg As String
    Dim found As Long

    ' If the report has never been built, build the whole report instead
    ' (RunFullReport also writes the anomaly list).
    If Not SheetExists(REPORT_SHEET) Then
        RunFullReport
        Exit Sub
    End If

    oldCalc = Application.Calculation
    On Error GoTo Fail

    Application.Calculate
    Application.ScreenUpdating = False

    LoadTableData
    found = WriteAnomalyList(GetOrCreateSheet(REPORT_SHEET))

CleanExit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    If errMsg = "" Then
        WriteLog "HighlightAnomalies", "OK - " & found & " anomalies listed"
        ThisWorkbook.Worksheets(REPORT_SHEET).Activate
        ThisWorkbook.Worksheets(REPORT_SHEET).Cells(ROW_ANOMALY, 1).Select
        MsgBox found & " unusual hour(s) found and listed on the report.", _
               vbInformation, "Highlight Anomalies"
    Else
        WriteLog "HighlightAnomalies", "ERROR - " & errMsg
        MsgBox "Could not list the anomalies." & vbNewLine & "Reason: " & errMsg, _
               vbExclamation, "Highlight Anomalies"
    End If
    Exit Sub

Fail:
    errMsg = Err.Description
    Resume CleanExit
End Sub


'==============================================================================
' 4. RESET FILTERS (table filters + Dashboard drop-downs)
'==============================================================================
Public Sub ResetFilters()
    Dim lo As ListObject
    Dim wsDash As Worksheet
    Dim wsAnom As Worksheet
    Dim errMsg As String

    On Error GoTo Fail

    ' (a) Clear any filter on the data table. ShowAllData only works when a
    '     filter is actually active, so we check FilterMode first.
    Set lo = ThisWorkbook.Worksheets(DATA_SHEET).ListObjects(DATA_TABLE)
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If

    ' (b) Clear the filter arrows on the Anomalies detail list as well.
    Set wsAnom = ThisWorkbook.Worksheets(ANOM_SHEET)
    If wsAnom.FilterMode Then wsAnom.ShowAllData

    ' (c) Put the Dashboard drop-downs back to "All" and the full date range.
    Set wsDash = ThisWorkbook.Worksheets(DASH_SHEET)
    wsDash.Range("B6").Value = "All"                      ' Day type
    wsDash.Range("H6").Value = "All"                      ' Weather
    wsDash.Range("N6").Value = "All"                      ' Month
    wsDash.Range("T6").Value = "All"                      ' Time band
    wsDash.Range("Z6").Value = DateSerial(2011, 1, 1)     ' From date
    wsDash.Range("AF6").Value = DateSerial(2011, 2, 14)   ' To date

    Application.Calculate
    WriteLog "ResetFilters", "OK - filters cleared, Dashboard set to All"
    wsDash.Activate
    MsgBox "All filters have been reset.", vbInformation, "Reset Filters"
    Exit Sub

Fail:
    errMsg = Err.Description
    WriteLog "ResetFilters", "ERROR - " & errMsg
    MsgBox "Could not reset the filters." & vbNewLine & "Reason: " & errMsg, _
           vbExclamation, "Reset Filters"
End Sub


'==============================================================================
' 5. BUILD THE REPORT (called by RunFullReport)
'==============================================================================
Private Sub BuildReport()
    Dim ws As Worksheet
    Dim r As Long, h As Long, d As Long, k As Long, i As Long
    Dim cnt As Double, cas As Double, reg As Double

    ' Totals
    Dim totCnt As Double, totCas As Double, totReg As Double

    ' Sums and counts per hour of day (0..23) and per weekday (0=Sun..6=Sat)
    Dim hrCnt(0 To 23) As Double, hrCas(0 To 23) As Double, hrReg(0 To 23) As Double
    Dim hrN(0 To 23) As Long
    Dim wdCnt(0 To 6) As Double, wdCas(0 To 6) As Double, wdReg(0 To 6) As Double
    Dim wdN(0 To 6) As Long

    ' Weather (code 1 = clear, 3 = light rain/snow) and weekday vs weekend
    Dim clearSum As Double, clearN As Long, rainSum As Double, rainN As Long
    Dim wkdaySum As Double, wkdayN As Long, wkendSum As Double, wkendN As Long
    Dim peakSum As Double                 ' weekday commute hours 7-9 and 16-19

    ' Results
    Dim peakHour As Long, quietHour As Long, bestDay As Long
    Dim rainImpact As Double, hasRain As Boolean
    Dim anomalyCount As Long, zLimit As Double
    Dim dayNames As Variant, dayOrder As Variant
    Dim tbl() As Variant                   ' array we write to the sheet in one go

    ' Data-check results
    Dim blankCells As Long, dupInstants As Long, sumMismatch As Long
    Dim sheetAnomalies As Long

    dayNames = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    dayOrder = Array(1, 2, 3, 4, 5, 6, 0)  ' show Monday first, Sunday last

    '--------------------------------------------------------------------
    ' A. One loop over all rows: add up totals, per hour and per weekday
    '--------------------------------------------------------------------
    For r = 1 To mRows
        cnt = NumValue(mData(r, colCnt))
        cas = NumValue(mData(r, colCasual))
        reg = NumValue(mData(r, colRegistered))
        h = CLng(NumValue(mData(r, colHr)))
        d = CLng(NumValue(mData(r, colWeekday)))

        totCnt = totCnt + cnt
        totCas = totCas + cas
        totReg = totReg + reg

        If h >= 0 And h <= 23 Then
            hrCnt(h) = hrCnt(h) + cnt
            hrCas(h) = hrCas(h) + cas
            hrReg(h) = hrReg(h) + reg
            hrN(h) = hrN(h) + 1
        End If

        If d >= 0 And d <= 6 Then
            wdCnt(d) = wdCnt(d) + cnt
            wdCas(d) = wdCas(d) + cas
            wdReg(d) = wdReg(d) + reg
            wdN(d) = wdN(d) + 1
        End If

        ' Weather: 1 = Clear, 3 = Light rain/snow
        Select Case CLng(NumValue(mData(r, colWeather)))
            Case 1
                clearSum = clearSum + cnt
                clearN = clearN + 1
            Case 3
                rainSum = rainSum + cnt
                rainN = rainN + 1
        End Select

        ' Weekday vs weekend, using our own DayType function
        If DayType(d) = "Weekend" Then
            wkendSum = wkendSum + cnt
            wkendN = wkendN + 1
        Else
            wkdaySum = wkdaySum + cnt
            wkdayN = wkdayN + 1
            ' Commute peak hours (same rule as the Peak_Hour column)
            If (h >= 7 And h <= 9) Or (h >= 16 And h <= 19) Then peakSum = peakSum + cnt
        End If
    Next r

    '--------------------------------------------------------------------
    ' B. Find the busiest and quietest hour and the busiest weekday
    '--------------------------------------------------------------------
    peakHour = 0
    quietHour = 0
    For h = 0 To 23
        If SafeAvg(hrCnt(h), hrN(h)) > SafeAvg(hrCnt(peakHour), hrN(peakHour)) Then peakHour = h
        If SafeAvg(hrCnt(h), hrN(h)) < SafeAvg(hrCnt(quietHour), hrN(quietHour)) Then quietHour = h
    Next h

    bestDay = 0
    For d = 0 To 6
        If SafeAvg(wdCnt(d), wdN(d)) > SafeAvg(wdCnt(bestDay), wdN(bestDay)) Then bestDay = d
    Next d

    ' Rain impact = (average in light rain / average in clear weather) - 1
    hasRain = (rainN > 0 And clearN > 0)
    If hasRain Then
        If SafeAvg(clearSum, clearN) > 0 Then
            rainImpact = SafeAvg(rainSum, rainN) / SafeAvg(clearSum, clearN) - 1
        Else
            hasRain = False
        End If
    End If

    zLimit = GetZThreshold()
    anomalyCount = CountAnomalies(zLimit)

    '--------------------------------------------------------------------
    ' C. Data checks (read-only: we count problems, we do not fix them)
    '--------------------------------------------------------------------
    blankCells = CountBlankBaseCells()
    dupInstants = CountDuplicateInstants()
    For r = 1 To mRows
        If Abs(NumValue(mData(r, colCnt)) - (NumValue(mData(r, colCasual)) + NumValue(mData(r, colRegistered)))) > 0.0001 Then
            sumMismatch = sumMismatch + 1
        End If
    Next r
    sheetAnomalies = CountAnomalyFlagsOnSheet()

    '--------------------------------------------------------------------
    ' D. Prepare the report sheet: clear old contents and old charts
    '--------------------------------------------------------------------
    Set ws = GetOrCreateSheet(REPORT_SHEET)
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B:G").ColumnWidth = 15

    ' Title and timestamp
    ws.Range("A1").Value = "Bike Sharing Demand - Automated Report"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Last updated: " & Format(Now, "dd-mmm-yyyy hh:nn") & _
                           "   |   Source: table " & DATA_TABLE & " (" & mRows & " hourly records, all rows, filters ignored)"
    ws.Range("A2").Font.Italic = True
    ws.Range("A3").Value = "Built by the RunFullReport macro. Do not type on this sheet: it is rebuilt every run."
    ws.Range("A3").Font.Color = RGB(118, 118, 118)

    '--------------------------------------------------------------------
    ' E. KPI block
    '--------------------------------------------------------------------
    SectionTitle ws, ROW_KPI, "Key numbers"
    PutKpi ws, ROW_KPI + 1, "Total rentals", totCnt, "#,##0", "Sum of cnt"
    PutKpi ws, ROW_KPI + 2, "Registered rentals", totReg, "#,##0", "Sum of registered"
    PutKpi ws, ROW_KPI + 3, "Casual rentals", totCas, "#,##0", "Sum of casual"
    PutKpi ws, ROW_KPI + 4, "Registered share", SafeAvg(totReg, totCnt), "0.0%", "Registered / total"
    PutKpi ws, ROW_KPI + 5, "Casual share", SafeAvg(totCas, totCnt), "0.0%", "Casual / total"
    PutKpi ws, ROW_KPI + 6, "Average rentals per hour", SafeAvg(totCnt, mRows), "0.0", "Total / number of hours"
    PutKpi ws, ROW_KPI + 7, "Peak hour (highest average)", HourLabel(peakHour), "@", _
           "avg " & Format(SafeAvg(hrCnt(peakHour), hrN(peakHour)), "0") & " rentals"
    PutKpi ws, ROW_KPI + 8, "Busiest weekday (highest average)", dayNames(bestDay), "@", _
           "avg " & Format(SafeAvg(wdCnt(bestDay), wdN(bestDay)), "0") & " rentals per hour"
    If hasRain Then
        PutKpi ws, ROW_KPI + 9, "Rain impact (light rain vs clear)", rainImpact, "+0%;-0%", "Change in average rentals per hour"
    Else
        PutKpi ws, ROW_KPI + 9, "Rain impact (light rain vs clear)", "n/a", "@", "No rainy or no clear hours"
    End If
    PutKpi ws, ROW_KPI + 10, "Anomalies (|z| >= " & zLimit & ")", anomalyCount, "0", "Hours far from the normal for that hour"

    '--------------------------------------------------------------------
    ' F. Data checks block
    '--------------------------------------------------------------------
    SectionTitle ws, ROW_CHECKS, "Data checks"
    PutCheck ws, ROW_CHECKS + 1, "Rows in the table", mRows, (mRows > 0)
    PutCheck ws, ROW_CHECKS + 2, "Blank or error cells in the 16 base columns", blankCells, (blankCells = 0)
    PutCheck ws, ROW_CHECKS + 3, "Duplicate instant values", dupInstants, (dupInstants = 0)
    PutCheck ws, ROW_CHECKS + 4, "Rows where cnt <> casual + registered", sumMismatch, (sumMismatch = 0)
    If sheetAnomalies >= 0 Then
        PutCheck ws, ROW_CHECKS + 5, "Anomalies sheet agrees with VBA count", sheetAnomalies, (sheetAnomalies = anomalyCount)
    Else
        ws.Cells(ROW_CHECKS + 5, 1).Value = "Anomalies sheet agrees with VBA count"
        ws.Cells(ROW_CHECKS + 5, 2).Value = "n/a"
        ws.Cells(ROW_CHECKS + 5, 3).Value = "Anomalies sheet not found"
    End If

    '--------------------------------------------------------------------
    ' G. Hourly summary table + line chart
    '--------------------------------------------------------------------
    SectionTitle ws, ROW_HOURLY, "Average rentals by hour of day"
    ws.Cells(ROW_HOURLY + 1, 1).Resize(1, 5).Value = Array("Hour", "Hours of data", "Avg casual", "Avg registered", "Avg total")
    HeaderStyle ws.Cells(ROW_HOURLY + 1, 1).Resize(1, 5)

    ReDim tbl(1 To 24, 1 To 5)
    For h = 0 To 23
        tbl(h + 1, 1) = h
        tbl(h + 1, 2) = hrN(h)
        tbl(h + 1, 3) = Round(SafeAvg(hrCas(h), hrN(h)), 1)
        tbl(h + 1, 4) = Round(SafeAvg(hrReg(h), hrN(h)), 1)
        tbl(h + 1, 5) = Round(SafeAvg(hrCnt(h), hrN(h)), 1)
    Next h
    ws.Cells(ROW_HOURLY + 2, 1).Resize(24, 5).Value = tbl
    ws.Cells(ROW_HOURLY + 2, 3).Resize(24, 3).NumberFormat = "0.0"
    ' Shade the peak hour row so it stands out
    ws.Cells(ROW_HOURLY + 2 + peakHour, 1).Resize(1, 5).Interior.Color = RGB(255, 235, 156)

    AddChart ws, "chtHourly", xlLine, _
             ws.Cells(ROW_HOURLY + 1, 3).Resize(25, 3), _
             ws.Cells(ROW_HOURLY + 2, 1).Resize(24, 1), _
             ws.Cells(ROW_HOURLY, 8), 520, 330, _
             "Average rentals by hour of day", "Hour of day", "Avg rentals"

    '--------------------------------------------------------------------
    ' H. Weekday summary table + column chart
    '--------------------------------------------------------------------
    SectionTitle ws, ROW_WEEKDAY, "Average rentals by day of week"
    ws.Cells(ROW_WEEKDAY + 1, 1).Resize(1, 5).Value = Array("Day", "Hours of data", "Avg casual", "Avg registered", "Avg total")
    HeaderStyle ws.Cells(ROW_WEEKDAY + 1, 1).Resize(1, 5)

    ReDim tbl(1 To 7, 1 To 5)
    For k = 0 To 6
        d = dayOrder(k)
        tbl(k + 1, 1) = dayNames(d)
        tbl(k + 1, 2) = wdN(d)
        tbl(k + 1, 3) = Round(SafeAvg(wdCas(d), wdN(d)), 1)
        tbl(k + 1, 4) = Round(SafeAvg(wdReg(d), wdN(d)), 1)
        tbl(k + 1, 5) = Round(SafeAvg(wdCnt(d), wdN(d)), 1)
        If d = bestDay Then ws.Cells(ROW_WEEKDAY + 2 + k, 1).Resize(1, 5).Interior.Color = RGB(255, 235, 156)
    Next k
    ws.Cells(ROW_WEEKDAY + 2, 1).Resize(7, 5).Value = tbl
    ws.Cells(ROW_WEEKDAY + 2, 3).Resize(7, 3).NumberFormat = "0.0"

    AddChart ws, "chtWeekday", xlColumnClustered, _
             ws.Cells(ROW_WEEKDAY + 1, 5).Resize(8, 1), _
             ws.Cells(ROW_WEEKDAY + 2, 1).Resize(7, 1), _
             ws.Cells(ROW_WEEKDAY, 8), 520, 260, _
             "Average rentals per hour by weekday", "Day", "Avg rentals"

    '--------------------------------------------------------------------
    ' I. Top 5 busiest individual hours
    '--------------------------------------------------------------------
    WriteTop5 ws

    '--------------------------------------------------------------------
    ' J. Insight sentences, written from the numbers above
    '--------------------------------------------------------------------
    SectionTitle ws, ROW_INSIGHTS, "What the numbers say (written automatically)"
    k = ROW_INSIGHTS + 1

    ws.Cells(k, 1).Value = "1. The busiest time is " & HourLabel(peakHour) & " with about " & _
        Format(SafeAvg(hrCnt(peakHour), hrN(peakHour)), "0") & " rentals per hour; the quietest is " & _
        HourLabel(quietHour) & " with about " & Format(SafeAvg(hrCnt(quietHour), hrN(quietHour)), "0") & "."
    k = k + 1

    ws.Cells(k, 1).Value = "2. Registered riders make " & Format(SafeAvg(totReg, totCnt), "0%") & _
        " of all " & Format(totCnt, "#,##0") & " rentals; casual riders only " & Format(SafeAvg(totCas, totCnt), "0%") & "."
    k = k + 1

    ws.Cells(k, 1).Value = "3. " & dayNames(bestDay) & " is the busiest day on average. Weekdays average " & _
        Format(SafeAvg(wkdaySum, wkdayN), "0") & " rentals per hour against " & _
        Format(SafeAvg(wkendSum, wkendN), "0") & " at weekends, which points to commuting."
    k = k + 1

    ws.Cells(k, 1).Value = "4. Weekday commute hours (7-9 am and 4-7 pm) carry " & _
        Format(SafeAvg(peakSum, totCnt), "0%") & " of all rentals."
    k = k + 1

    If hasRain Then
        If rainImpact < 0 Then
            ws.Cells(k, 1).Value = "5. In light rain or snow, rentals per hour are " & Format(-rainImpact, "0%") & _
                " lower than in clear weather."
        Else
            ws.Cells(k, 1).Value = "5. In light rain or snow, rentals per hour are " & Format(rainImpact, "0%") & _
                " higher than in clear weather (unusual - check the data)."
        End If
    Else
        ws.Cells(k, 1).Value = "5. Rain impact could not be measured (no rainy or no clear hours)."
    End If
    k = k + 1

    ws.Cells(k, 1).Value = "6. " & anomalyCount & " hour(s) were unusual (|z| >= " & zLimit & _
        " compared with the normal for the same hour). They are listed at the bottom of this sheet."
    k = k + 1

    If blankCells = 0 And dupInstants = 0 And sumMismatch = 0 Then
        ws.Cells(k, 1).Value = "7. All data checks passed: no blanks, no duplicate instants, and cnt = casual + registered on every row."
    Else
        ws.Cells(k, 1).Value = "7. Some data checks need attention: see the Data checks block above (cells marked 'Check')."
        ws.Cells(k, 1).Font.Color = RGB(192, 0, 0)
    End If
End Sub


'==============================================================================
' 6. ANOMALY LIST (used by RunFullReport and HighlightAnomalies)
'    Returns how many anomalies were written.
'==============================================================================
Private Function WriteAnomalyList(ByVal ws As Worksheet) As Long
    Dim meanH() As Double, sdH() As Double, nH() As Long
    Dim r As Long, h As Long, outRow As Long, found As Long
    Dim cnt As Double, z As Double, zLimit As Double

    zLimit = GetZThreshold()
    GetHourStats meanH, sdH, nH

    ' Clear only the anomaly area (columns A to G, from ROW_ANOMALY down)
    ws.Range(ws.Cells(ROW_ANOMALY, 1), ws.Cells(ws.Rows.Count, 7)).Clear

    SectionTitle ws, ROW_ANOMALY, "Anomalies: hours with |z| >= " & zLimit & " (orange = busier than normal, blue = quieter)"
    ws.Cells(ROW_ANOMALY + 1, 1).Resize(1, 7).Value = _
        Array("Instant", "Date", "Hour", "Rentals (cnt)", "Normal for this hour", "Z-score", "Flag")
    HeaderStyle ws.Cells(ROW_ANOMALY + 1, 1).Resize(1, 7)

    outRow = ROW_ANOMALY + 2
    For r = 1 To mRows
        h = CLng(NumValue(mData(r, colHr)))
        cnt = NumValue(mData(r, colCnt))
        If h >= 0 And h <= 23 Then
            ' z-score = how many standard deviations away from the hour's mean
            If sdH(h) > 0 Then z = (cnt - meanH(h)) / sdH(h) Else z = 0
            If Abs(z) >= zLimit Then
                ws.Cells(outRow, 1).Value = mData(r, colInstant)
                ws.Cells(outRow, 2).Value = mData(r, colDteday)
                ws.Cells(outRow, 2).NumberFormat = "dd-mmm-yyyy"
                ws.Cells(outRow, 3).Value = HourLabel(h)
                ws.Cells(outRow, 4).Value = cnt
                ws.Cells(outRow, 5).Value = Round(meanH(h), 1)
                ws.Cells(outRow, 6).Value = Round(z, 2)
                If z > 0 Then
                    ws.Cells(outRow, 7).Value = "High"
                    ws.Cells(outRow, 1).Resize(1, 7).Interior.Color = RGB(252, 213, 180)   ' orange
                Else
                    ws.Cells(outRow, 7).Value = "Low"
                    ws.Cells(outRow, 1).Resize(1, 7).Interior.Color = RGB(189, 215, 238)   ' blue
                End If
                outRow = outRow + 1
                found = found + 1
            End If
        End If
    Next r

    If found = 0 Then ws.Cells(outRow, 1).Value = "No anomalies at this threshold."
    WriteAnomalyList = found
End Function


'==============================================================================
' 7. TOP 5 BUSIEST HOURS
'    Simple method: 5 times, scan all rows and pick the biggest cnt that has
'    not been picked yet.
'==============================================================================
Private Sub WriteTop5(ByVal ws As Worksheet)
    Dim used() As Boolean
    Dim meanH() As Double, sdH() As Double, nH() As Long
    Dim rank As Long, r As Long, best As Long, h As Long
    Dim rowOut As Long

    ReDim used(1 To mRows)
    GetHourStats meanH, sdH, nH

    SectionTitle ws, ROW_TOP5, "Top 5 busiest hours"
    ws.Cells(ROW_TOP5 + 1, 1).Resize(1, 7).Value = _
        Array("Rank", "Date", "Hour", "Day", "Weather", "Rentals (cnt)", "Vs normal for hour")
    HeaderStyle ws.Cells(ROW_TOP5 + 1, 1).Resize(1, 7)

    For rank = 1 To 5
        If rank > mRows Then Exit For
        best = 0
        For r = 1 To mRows
            If Not used(r) Then
                If best = 0 Then
                    best = r
                ElseIf NumValue(mData(r, colCnt)) > NumValue(mData(best, colCnt)) Then
                    best = r
                End If
            End If
        Next r
        used(best) = True

        rowOut = ROW_TOP5 + 1 + rank
        h = CLng(NumValue(mData(best, colHr)))
        ws.Cells(rowOut, 1).Value = rank
        ws.Cells(rowOut, 2).Value = mData(best, colDteday)
        ws.Cells(rowOut, 2).NumberFormat = "dd-mmm-yyyy"
        ws.Cells(rowOut, 3).Value = HourLabel(h)
        ws.Cells(rowOut, 4).Value = DayName(mData(best, colWeekday)) & " (" & DayType(mData(best, colWeekday)) & ")"
        ws.Cells(rowOut, 5).Value = WeatherLabel(mData(best, colWeather))
        ws.Cells(rowOut, 6).Value = NumValue(mData(best, colCnt))
        If h >= 0 And h <= 23 Then
            ws.Cells(rowOut, 7).Value = DemandLevel(NumValue(mData(best, colCnt)), meanH(h), sdH(h))
        End If
    Next rank
End Sub


'==============================================================================
' 8. CUSTOM FUNCTIONS (UDFs) - also usable in worksheet cells
'==============================================================================

' Weekday code 0 = Sunday ... 6 = Saturday. Returns "Weekend" or "Weekday".
Public Function DayType(ByVal weekdayCode As Variant) As String
    If IsError(weekdayCode) Then
        DayType = "Unknown"
    ElseIf Not IsNumeric(weekdayCode) Then
        DayType = "Unknown"
    Else
        Select Case CLng(weekdayCode)
            Case 0, 6
                DayType = "Weekend"
            Case 1 To 5
                DayType = "Weekday"
            Case Else
                DayType = "Unknown"
        End Select
    End If
End Function

' Weather code 1-4 from the UCI data dictionary -> readable label.
Public Function WeatherLabel(ByVal code As Variant) As String
    If IsError(code) Then
        WeatherLabel = "Unknown"
    ElseIf Not IsNumeric(code) Then
        WeatherLabel = "Unknown"
    Else
        Select Case CLng(code)
            Case 1: WeatherLabel = "Clear"
            Case 2: WeatherLabel = "Cloudy/Mist"
            Case 3: WeatherLabel = "Light Rain/Snow"
            Case 4: WeatherLabel = "Heavy Rain"
            Case Else: WeatherLabel = "Unknown"
        End Select
    End If
End Function

' Compares a rental count with the typical demand for that hour:
' more than 1 standard deviation above the mean = "High",
' more than 1 standard deviation below = "Low", otherwise "Normal".
Public Function DemandLevel(ByVal rentals As Double, ByVal hourMean As Double, ByVal hourSd As Double) As String
    If hourSd <= 0 Then
        DemandLevel = "Normal"
    ElseIf rentals >= hourMean + hourSd Then
        DemandLevel = "High"
    ElseIf rentals <= hourMean - hourSd Then
        DemandLevel = "Low"
    Else
        DemandLevel = "Normal"
    End If
End Function


'==============================================================================
' 9. HELPERS (small building blocks used above)
'==============================================================================

' Reads tblFinal into the array mData and finds the column numbers.
Private Sub LoadTableData()
    Dim lo As ListObject

    If Not SheetExists(DATA_SHEET) Then
        Err.Raise vbObjectError + 513, , "Sheet '" & DATA_SHEET & "' was not found."
    End If
    Set lo = ThisWorkbook.Worksheets(DATA_SHEET).ListObjects(DATA_TABLE)
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 514, , "Table '" & DATA_TABLE & "' has no data rows."
    End If

    ' .Value of a range gives a 2-D array (rows, columns) of the RESULTS of
    ' the formulas - we never touch the formulas themselves.
    mData = lo.DataBodyRange.Value
    mRows = UBound(mData, 1)

    colInstant = ColumnNumber(lo, "instant")
    colDteday = ColumnNumber(lo, "dteday")
    colHr = ColumnNumber(lo, "hr")
    colWeekday = ColumnNumber(lo, "weekday")
    colWeather = ColumnNumber(lo, "weathersit")
    colCasual = ColumnNumber(lo, "casual")
    colRegistered = ColumnNumber(lo, "registered")
    colCnt = ColumnNumber(lo, "cnt")
End Sub

' Finds a column by its header name; gives a clear message if it is missing.
Private Function ColumnNumber(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If LCase$(Trim$(lc.Name)) = LCase$(headerName) Then
            ColumnNumber = lc.Index
            Exit Function
        End If
    Next lc
    Err.Raise vbObjectError + 515, , "Column '" & headerName & "' was not found in table " & lo.Name & "."
End Function

' Mean and sample standard deviation of cnt for each hour 0..23
' (the same maths as the Anomalies sheet: STDEV with n - 1).
Private Sub GetHourStats(ByRef meanH() As Double, ByRef sdH() As Double, ByRef nH() As Long)
    Dim sumH(0 To 23) As Double, sqDev(0 To 23) As Double
    Dim r As Long, h As Long, cnt As Double

    ReDim meanH(0 To 23)
    ReDim sdH(0 To 23)
    ReDim nH(0 To 23)

    ' Pass 1: add up cnt per hour, then divide to get the mean
    For r = 1 To mRows
        h = CLng(NumValue(mData(r, colHr)))
        If h >= 0 And h <= 23 Then
            sumH(h) = sumH(h) + NumValue(mData(r, colCnt))
            nH(h) = nH(h) + 1
        End If
    Next r
    For h = 0 To 23
        If nH(h) > 0 Then meanH(h) = sumH(h) / nH(h)
    Next h

    ' Pass 2: add up the squared distance from the mean, then take the root
    For r = 1 To mRows
        h = CLng(NumValue(mData(r, colHr)))
        If h >= 0 And h <= 23 Then
            cnt = NumValue(mData(r, colCnt))
            sqDev(h) = sqDev(h) + (cnt - meanH(h)) ^ 2
        End If
    Next r
    For h = 0 To 23
        If nH(h) > 1 Then sdH(h) = Sqr(sqDev(h) / (nH(h) - 1))
    Next h
End Sub

' Counts rows whose |z| is at or above the threshold.
Private Function CountAnomalies(ByVal zLimit As Double) As Long
    Dim meanH() As Double, sdH() As Double, nH() As Long
    Dim r As Long, h As Long, z As Double, total As Long

    GetHourStats meanH, sdH, nH
    For r = 1 To mRows
        h = CLng(NumValue(mData(r, colHr)))
        If h >= 0 And h <= 23 Then
            If sdH(h) > 0 Then
                z = (NumValue(mData(r, colCnt)) - meanH(h)) / sdH(h)
                If Abs(z) >= zLimit Then total = total + 1
            End If
        End If
    Next r
    CountAnomalies = total
End Function

' Reads the z threshold from Anomalies!B4 (uses 3 if the cell is not usable).
Private Function GetZThreshold() As Double
    Dim v As Variant
    GetZThreshold = 3
    On Error Resume Next
    v = ThisWorkbook.Worksheets(ANOM_SHEET).Range(ANOM_THRESHOLD_CELL).Value
    If Not IsError(v) Then
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CDbl(v) > 0 Then GetZThreshold = CDbl(v)
        End If
    End If
End Function

' Counts "High"/"Low" flags already shown on the Anomalies sheet, so we can
' check that the VBA result and the formula result agree. -1 = not available.
Private Function CountAnomalyFlagsOnSheet() As Long
    Dim v As Variant, r As Long, total As Long
    CountAnomalyFlagsOnSheet = -1
    If Not SheetExists(ANOM_SHEET) Then Exit Function
    On Error GoTo NotAvailable
    v = ThisWorkbook.Worksheets(ANOM_SHEET).Range(ANOM_FLAG_RANGE).Value
    For r = 1 To UBound(v, 1)
        If Not IsError(v(r, 1)) Then
            If Len(CStr(v(r, 1))) > 0 Then total = total + 1
        End If
    Next r
    CountAnomalyFlagsOnSheet = total
NotAvailable:
End Function

' Counts empty or error cells in the 16 base columns.
Private Function CountBlankBaseCells() As Long
    Dim lo As ListObject, baseNames As Variant
    Dim i As Long, c As Long, r As Long, total As Long
    Dim v As Variant

    Set lo = ThisWorkbook.Worksheets(DATA_SHEET).ListObjects(DATA_TABLE)
    baseNames = Split(BASE_COLUMNS, ",")
    For i = LBound(baseNames) To UBound(baseNames)
        c = ColumnNumber(lo, CStr(baseNames(i)))
        For r = 1 To mRows
            v = mData(r, c)
            If IsError(v) Then
                total = total + 1
            ElseIf IsEmpty(v) Then
                total = total + 1
            ElseIf Trim$(CStr(v)) = "" Then
                total = total + 1
            End If
        Next r
    Next i
    CountBlankBaseCells = total
End Function

' Counts repeated instant values. A Collection only accepts each key once,
' so adding a key that already exists raises an error = a duplicate.
' (A Collection works on Mac; Scripting.Dictionary does not.)
Private Function CountDuplicateInstants() As Long
    Dim seen As Collection, r As Long, total As Long
    Dim key As String

    Set seen = New Collection
    On Error Resume Next
    For r = 1 To mRows
        If Not IsError(mData(r, colInstant)) Then
            key = "k" & CStr(mData(r, colInstant))
            Err.Clear
            seen.Add r, key
            If Err.Number <> 0 Then total = total + 1
        End If
    Next r
    Err.Clear
    On Error GoTo 0
    CountDuplicateInstants = total
End Function

' Turns any cell value into a number. Errors and text become 0.
Private Function NumValue(ByVal v As Variant) As Double
    If IsError(v) Then
        NumValue = 0
    ElseIf IsEmpty(v) Then
        NumValue = 0
    ElseIf IsNumeric(v) Then
        NumValue = CDbl(v)
    Else
        NumValue = 0
    End If
End Function

' Division that returns 0 instead of a "divide by zero" error.
Private Function SafeAvg(ByVal total As Double, ByVal count As Double) As Double
    If count = 0 Then
        SafeAvg = 0
    Else
        SafeAvg = total / count
    End If
End Function

' Weekday code 0..6 -> "Sun".."Sat"
Private Function DayName(ByVal weekdayCode As Variant) As String
    Dim dayNames As Variant
    dayNames = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    DayName = "?"
    If IsError(weekdayCode) Then Exit Function
    If Not IsNumeric(weekdayCode) Then Exit Function
    If CLng(weekdayCode) >= 0 And CLng(weekdayCode) <= 6 Then DayName = dayNames(CLng(weekdayCode))
End Function

' 8 -> "08:00"
Private Function HourLabel(ByVal h As Long) As String
    HourLabel = Format(h, "00") & ":00"
End Function

' Writes one KPI row: label, value, number format and a short note.
Private Sub PutKpi(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                   ByVal v As Variant, ByVal numFmt As String, ByVal note As String)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).NumberFormat = numFmt
    ws.Cells(r, 2).Value = v
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).HorizontalAlignment = xlRight
    ws.Cells(r, 3).Value = note
    ws.Cells(r, 3).Font.Color = RGB(118, 118, 118)
End Sub

' Writes one data-check row: label, number found and OK / Check.
Private Sub PutCheck(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                     ByVal v As Variant, ByVal passed As Boolean)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = v
    ws.Cells(r, 2).HorizontalAlignment = xlRight
    If passed Then
        ws.Cells(r, 3).Value = "OK"
        ws.Cells(r, 3).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(r, 3).Value = "Check"
        ws.Cells(r, 3).Font.Color = RGB(192, 0, 0)
    End If
    ws.Cells(r, 3).Font.Bold = True
End Sub

' Bold blue section heading.
Private Sub SectionTitle(ByVal ws As Worksheet, ByVal r As Long, ByVal text As String)
    ws.Cells(r, 1).Value = text
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 13
    ws.Cells(r, 1).Font.Color = RGB(31, 78, 121)
End Sub

' Dark header row with white bold text.
Private Sub HeaderStyle(ByVal rng As Range)
    rng.Font.Bold = True
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = RGB(31, 78, 121)
    rng.HorizontalAlignment = xlCenter
End Sub

' Creates a chart from a range.
'   seriesRange : the data columns INCLUDING the header row (header = series name)
'   labelRange  : the category labels for the x-axis (hours or day names)
'   anchor      : the cell where the top-left corner of the chart goes
Private Sub AddChart(ByVal ws As Worksheet, ByVal chartName As String, ByVal chartKind As Long, _
                     ByVal seriesRange As Range, ByVal labelRange As Range, ByVal anchor As Range, _
                     ByVal w As Double, ByVal ht As Double, ByVal titleText As String, _
                     ByVal xTitle As String, ByVal yTitle As String)
    Dim co As ChartObject
    Dim i As Long

    Set co = ws.ChartObjects.Add(Left:=anchor.Left, Top:=anchor.Top, Width:=w, Height:=ht)
    co.Name = chartName
    With co.Chart
        ' Give the chart its data first, then choose the chart type
        .SetSourceData Source:=seriesRange, PlotBy:=xlColumns
        .ChartType = chartKind
        ' Use our labels (hour / day) for the x-axis of every series
        For i = 1 To .SeriesCollection.Count
            .SeriesCollection(i).XValues = labelRange
        Next i
        .HasTitle = True
        .ChartTitle.Text = titleText
        .HasLegend = (.SeriesCollection.Count > 1)
        If .HasLegend Then .Legend.Position = xlLegendPositionBottom
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = xTitle
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = yTitle
    End With
End Sub

' Refreshes data connections if there are any. Errors are ignored because
' this workbook normally has no external connections.
Private Sub RefreshConnectionsSafely()
    On Error Resume Next
    If ThisWorkbook.Connections.Count > 0 Then ThisWorkbook.RefreshAll
    On Error GoTo 0
End Sub

' True if a sheet with this name exists in the workbook.
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

' Returns the sheet, creating it at the end of the workbook if needed.
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    If SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

' Adds one line to the "Automation Log" sheet: time, macro name, result.
' It never stops the main macro, even if logging fails.
Private Sub WriteLog(ByVal macroName As String, ByVal result As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = GetOrCreateSheet(LOG_SHEET)
    If ws Is Nothing Then Exit Sub
    If ws.Range("A1").Value = "" Then
        ws.Range("A1:C1").Value = Array("Timestamp", "Macro", "Result")
        ws.Range("A1:C1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 20
        ws.Columns("B").ColumnWidth = 20
        ws.Columns("C").ColumnWidth = 80
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    ws.Cells(r, 2).Value = macroName
    ws.Cells(r, 3).Value = result
End Sub
