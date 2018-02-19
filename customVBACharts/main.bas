Attribute VB_Name = "main"

Sub updateRMCC()
    ' creates an object for the charts update properties. Give this a unique name.
    Dim rmccChart As New customChart
    ' in setProperties, arguments are: Name of Sheet where Graphs live,
    ' Name of Sheet where Data Lives, Column in data sheet for horizontal axis,
    ' Cell on Graph Sheet where Series identifier will be typed.
    If rmccChart.setProperties("Manual", "RMCC DEMAND", 1, "$D$2") Then
    ' Give updateSeries the value of the charts index. If its the first chart on the sheet 1, second =2, etc.
    ' Note this value is based off of when the chart was added, not its position in the sheet.
    rmccChart.updateSeries (1)
    End If
    
End Sub

Sub updateTargetInventory()

    Dim targetChart As New customChart
    If targetChart.setProperties("Manual", "Target Inventory", 7, "$D$2") Then
    targetChart.updateSeries (2)
    End If
End Sub

Sub test()

Debug.Print Worksheets("Manual").TextBox1.Text

End Sub
