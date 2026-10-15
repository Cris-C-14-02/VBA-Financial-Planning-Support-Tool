Attribute VB_Name = "FinanceCharts"
Option Explicit

Private Function ReportChart(ByVal chartName As String, ByVal left As Double) As ChartObject
    Dim ws As Worksheet, existing As ChartObject
    Set ws = AppSheet("Output")
    On Error Resume Next
    Set existing = ws.ChartObjects(chartName)
    On Error GoTo 0
    If existing Is Nothing Then Set existing = ws.ChartObjects.Add(left, ws.Cells(Application.Max(12, LastDataRow(ws, 5, 10) + 3), 1).Top, 460, 270)
    existing.Name = chartName
    existing.left = left: existing.Top = ws.Cells(Application.Max(12, LastDataRow(ws, 5, 10) + 3), 1).Top
    existing.Width = 460: existing.Height = 270
    Set ReportChart = existing
End Function

Public Sub UpdateSavingsChart(ByVal lastRow As Long)
    Dim ws As Worksheet, chart As chart
    Set ws = AppSheet("Output")
    Set chart = ReportChart("SavingsBalance", ws.Range("A1").left).chart
    Do While chart.SeriesCollection.Count > 0: chart.SeriesCollection(1).Delete: Loop
    chart.ChartType = xlLineMarkers
    With chart.SeriesCollection.NewSeries
        .Name = "Savings balance"
        .XValues = ws.Range("E2:E" & lastRow)
        .Values = ws.Range("J2:J" & lastRow)
        .Format.Line.ForeColor.RGB = RGB(13, 148, 136)
    End With
    chart.HasTitle = True: chart.ChartTitle.Text = "Savings balance over time"
    chart.HasLegend = False
    chart.Axes(xlCategory).CategoryType = xlCategoryScale
    chart.Axes(xlCategory).TickLabels.NumberFormat = "mmm d"
    chart.Axes(xlValue).TickLabels.NumberFormat = "$#,##0"
End Sub

Public Sub UpdateSpendingChart()
    Dim ws As Worksheet, chart As chart
    Set ws = AppSheet("Output")
    Set chart = ReportChart("SpendingBreakdown", ws.Range("A1").left + 480).chart
    Do While chart.SeriesCollection.Count > 0: chart.SeriesCollection(1).Delete: Loop
    chart.ChartType = xlDoughnut
    With chart.SeriesCollection.NewSeries
        .Name = "Spending"
        .XValues = ws.Range("L2:L5")
        .Values = ws.Range("M2:M5")
        .ApplyDataLabels
        .DataLabels.ShowValue = False
        .DataLabels.ShowPercentage = True
        .DataLabels.NumberFormat = "0.0%"
        .Points(1).Format.Fill.ForeColor.RGB = RGB(13, 148, 136)
        .Points(2).Format.Fill.ForeColor.RGB = RGB(59, 130, 246)
        .Points(3).Format.Fill.ForeColor.RGB = RGB(245, 158, 11)
        .Points(4).Format.Fill.ForeColor.RGB = RGB(148, 163, 184)
    End With
    chart.HasTitle = True: chart.ChartTitle.Text = "Spending by category"
    chart.HasLegend = True: chart.Legend.Position = xlLegendPositionBottom
End Sub
