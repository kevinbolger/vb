VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "customChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private trendingID As String
Private trendingCol As Integer
Private trendingRange As Range
Private trendingSheet As String
Private frontEndSheet As String
Private lRow As Long
Private controlRange As String
Private categoryCol As Integer
Private xAxisRange As Range


Private Sub setControlRange(cr As String)

    controlRange = cr

End Sub

Private Sub setCategoryCol(colNum As Integer)

categoryCol = colNum

End Sub

Private Sub setLastRow()

    'Find the last used row in a Column: column A in this example
    Dim LastRow As Long
    With Worksheets(trendingSheet)
        lRow = .Cells(.Rows.Count, trendingCol).End(xlUp).Row
    End With
    Exit Sub

End Sub

Private Sub setTrendingSheet(theSheet As String)

    trendingSheet = theSheet
    Exit Sub

End Sub

Private Sub setFrontEndSheet(sheetName As String)

    frontEndSheet = sheetName
    Exit Sub

    
End Sub


Private Sub setTrendingID(theSheet As String)

    trendingID = Worksheets(theSheet).Range(controlRange).Value

End Sub

Private Sub setTrendingCol()

    'Below solution inspired by following thread on stackoverflow:
    'https://stackoverflow.com/questions/13769375/excel-vba-return-first-occurrence-of-a-word-in-a-column-the-last
    trendingCol = Worksheets(trendingSheet).Rows(1).Find(What:=trendingID, LookAt:=xlWhole, MatchCase:=False).Column

End Sub


Private Sub setTrendingRange()

    Set c1 = Worksheets(trendingSheet).Cells(2, trendingCol)
    Set c2 = Worksheets(trendingSheet).Cells(lRow, trendingCol)
    r = c1.Address() & ":" & c2.Address()
    Set trendingRange = Worksheets(trendingSheet).Range(r)

End Sub

Private Sub setxAxisRange()

    Set c1 = Worksheets(trendingSheet).Cells(2, categoryCol)
    Set c2 = Worksheets(trendingSheet).Cells(lRow, categoryCol)
    r = c1.Address() & ":" & c2.Address()
    Set xAxisRange = Worksheets(trendingSheet).Range(r)

End Sub

Public Function setProperties(frontEndSheet As String, dataSheet As String, catCol As Integer, cr As String) As Boolean

    On Error GoTo eh
    setControlRange cr
    setTrendingID frontEndSheet
    setTrendingSheet dataSheet
    setFrontEndSheet frontEndSheet
    setTrendingCol
    setLastRow
    setCategoryCol catCol
    setxAxisRange
    setTrendingRange
    setProperties = True
    Exit Function
eh:
    setProperties = False
    
End Function


Public Sub updateSeries(chartObjectNum As Integer)

    Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.HasTitle = True
    If Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.SeriesCollection.Count > 0 Then
        Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.SeriesCollection(1).Delete
    End If
    Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.SeriesCollection.Add Source:=Worksheets(trendingSheet).Range(trendingRange.Address)
    Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.SeriesCollection(1).Name = trendingID
    Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.SeriesCollection(1).Trendlines.Add
    Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.HasTitle = True
    Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.ChartTitle.Text = trendingID
    Worksheets(frontEndSheet).ChartObjects(chartObjectNum).Chart.SeriesCollection(1).XValues = xAxisRange

End Sub

