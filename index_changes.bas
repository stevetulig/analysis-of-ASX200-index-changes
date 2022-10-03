Attribute VB_Name = "Module1"
Public Sub RefreshChart()

    Dim d_change As Date, d_analysis As Date
    Dim strdc As String, strda As String
    d_change = Range("change_date_selected").Value
    d_analysis = Range("analysis_date_selected").Value
    strdc = Year(d_change) & Format(Month(d_change), "00") & Format(Day(d_change), "00")
    strda = Year(d_analysis) & Format(Month(d_analysis), "00") & Format(Day(d_analysis), "00")

    Call GetExecIndexChangeData(strdc, strda)
    DrawScatterPlot

End Sub
Private Sub GetExecIndexChangeData(strdc As String, strda As String)

    Dim cnxn As ADODB.Connection
    Dim strConn As String
    Dim cmd As ADODB.Command
    Dim prm As ADODB.Parameter
    Dim rs As ADODB.Recordset
    
    Dim ws As Worksheet
    
    Set ws = Sheets("Data")
    
    Set cnxn = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rs = New ADODB.Recordset
    
    strConn = "PROVIDER=SQLOLEDB; "
    strConn = strConn & "DATA SOURCE=STEVE_XPS\SQLEXPRESS; INITIAL CATALOG=Zenith; "
    strConn = strConn & "INTEGRATED SECURITY=sspi;"
    
    cnxn.Open strConn
    With rs
        .ActiveConnection = cnxn
        .Open "exec index_change_analysis" & " '" & strdc & "', '" & strda & "'"
    End With
    ws.Range("A:F").ClearContents
    With rs
        For i = 1 To .Fields.Count
            ws.Cells(1, i) = .Fields(i - 1).Name
        Next i
    End With
    ws.Range("A2").CopyFromRecordset rs
    rs.Close
    
    cnxn.Close
    
End Sub
Private Sub DrawScatterPlot()

    Dim ws As Worksheet
    Set ws = Sheets("Main")

    ws.ChartObjects("Chart 1").Activate
    With ActiveChart
    
        For i = 1 To .FullSeriesCollection.Count
            .FullSeriesCollection(1).Delete
        Next i
        
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(1)
            .Name = "Constituent"
            .XValues = "=Data!$B$2:$B$5000"
            .Values = "=Data!$C$2:$C$5000"
            .MarkerForegroundColorIndex = 4
            .MarkerBackgroundColorIndex = 4
            .MarkerSize = 3
        End With
    
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(2)
            .Name = "Addition"
            .XValues = "=Data!$B$2:$B$5000"
            .Values = "=Data!$D$2:$D$5000"
            .MarkerForegroundColorIndex = 5
            .MarkerBackgroundColorIndex = 5
        End With
        
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(3)
            .Name = "Removal"
            .XValues = "=Data!$B$2:$B$5000"
            .Values = "=Data!$E$2:$E$5000"
            .MarkerForegroundColorIndex = 3
            .MarkerBackgroundColorIndex = 3
        End With
    
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(4)
            .Name = "Other"
            .XValues = "=Data!$B$2:$B$5000"
            .Values = "=Data!$F$2:$F$5000"
            .MarkerStyle = -4168
            .MarkerSize = 3
            .MarkerForegroundColorIndex = 1
        End With
    
        .Axes(xlCategory).MaximumScale = Range("XLIM").Value
        .Axes(xlValue).MaximumScale = Range("YLIM").Value
        .Axes(xlValue).MinimumScale = 0
        .HasTitle = True
        .ChartTitle.Caption = ChartTitle()
    End With

End Sub
Private Function ChartTitle() As String

    ChartTitle = "Analysis of ASX200 index changes on " & Range("change_date_selected").Value
    ChartTitle = ChartTitle & " Rankings by liquidity and market cap on " & Range("analysis_date_selected").Value

End Function
