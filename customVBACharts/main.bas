Attribute VB_Name = "main"

'The following is an example of the chart being implemented to update a chart.
Sub updateSomeChart()
    ' creates an object for the charts update properties. Give this a unique name.
    Dim someChart As New customChart
    ' in setProperties, arguments are: Name of Sheet where Graphs live,
    ' Name of Sheet where Data Lives, Column in data sheet for horizontal axis,
    ' Cell on Graph Sheet where Series identifier will be typed.
    If someChart.setProperties("Manual", "RMCC DEMAND", 1, "$D$2") Then
    ' Give updateSeries the value of the charts index. If its the first chart on the sheet 1, second =2, etc.
    ' Note this value is based off of when the chart was added, not its position in the sheet.
    someChart.updateSeries (1)
    End If
    
End Sub

' This class can be reused on any chart in any sheet. The only requirements are:

'1) The data are in a fat, short format (each data series has its own column)
'2) The column specified for the horizontal axis is static for a given chart
'3) The control field and chart are on the same sheet
