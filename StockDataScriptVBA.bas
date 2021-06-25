Attribute VB_Name = "Module1"
Sub stonks()

Dim headers() As Variant
Dim MainWs As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook

'Store header variables
headers() = Array("Ticker ", "Date ", "Open", "High", "Low", "Close", "Volume", " ", _
"Ticker", "Yearly_Change", "Percent_Change", "Stock_Volume", " ", " ", " ", "Ticker", "Value")

'For loop to run script in each Worksheet

For Each MainWs In wb.Sheets
    With MainWs
    .Rows(1).Value = ""
    For i = LBound(headers()) To UBound(headers())
    .Cells(1, 1 + i).Value = headers(i)
    
    Next i
    .Rows(1).Font.Bold = True
    .Rows(1).VerticalAlignment = xlCenter
    End With
Next MainWs

    For Each MainWs In Worksheets
    
    'Dimension and Initialize Variables
    
        Dim TickerName As String
        Dim TotalTickerVolume As Double
        Dim BegPrice As Double
        Dim EndPrice As Double
        Dim YearlyPriceChange As Double
        Dim YearlyPriceChangePercent As Double
        Dim MaxTickerName As String
        Dim MinTickerName As String
        Dim MaxPercent As Double
        Dim MaxVolumeTickerName As String
        Dim MaxVolume As Double

        TickerName = " "
        TotalTickerVolume = 0
        BegPrice = 0
        EndPrice = 0
        YearlyPriceChange = 0
        YearlyPriceChangePercent = 0
        MaxTickerName = " "
        MinTickerName = " "
        MaxPercent = 0
        MaxVolumeTickerName = " "
        MaxVolume = 0
        
        'location for variables in the summary table
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        
        'Create LastRow variable for loops to go to last non empty cell
        
        
        Dim LastRow As Long
    
        LastRow = MainWs.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Store beginning stock value for first ticker
        BegPrice = MainWs.Cells(2, 3).Value
        
       'Loop from top of main worksheet till last row in the last worksheet
        For i = 2 To LastRow
      
      'Check Ticker name
      
            If MainWs.Cells(i + 1, 1).Value <> MainWs.Cells(i, 1).Value Then
        
        'Set ticker name starting point
        
                TickerName = MainWs.Cells(i, 1).Value
            
                
                'Calculate Price Change
                EndPrice = MainWs.Cells(i, 6).Value
                YearlyPriceChange = EndPrice - BegPrice
                
                'Set condition for no change
                If BegPrice <> 0 Then
                     YearlyPriceChangePercent = (YearlyPriceChange / BegPrice) * 100
            
                End If
        
        'Add to ticker name the total volume
        TotalTickerVolume = TotalTickerVolume + MainWs.Cells(i, 7).Value
        
    
        'print the ticker name in the summary table
        MainWs.Cells(SummaryTableRow, 9).Value = TickerName
        
        'print the yearly price change in the summary table
        MainWs.Cells(SummaryTableRow, 10).Value = YearlyPriceChange
        
        'color fill yearly price change
        
        If (YearlyPriceChange > 0) Then
            MainWs.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
            
        ElseIf (YearlyPriceChange <= 0) Then
            
            MainWs.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
        End If
        
        'Print Yearly price change as a percent in summary table
        MainWs.Cells(SummaryTableRow, 11).Value = (CStr(YearlyPriceChangePercent) & "%")
        
        'Print total stock volume in summary table
        MainWs.Cells(SummaryTableRow, 12).Value = TotalTickerVolume
        
        ' add 1 to the summary table row count
        SummaryTableRow = SummaryTableRow + 1
        
        'Go to next beginning price
        BegPrice = MainWs.Cells(i + 1, 3).Value
        
        If (YearlyPriceChangePercent > MaxPercent) Then
            MaxPercent = YearlyPriceChangePercent
            MaxTickerName = TickerName
        
        ElseIf (YearlyPriceChangePercent < MinPercent) Then
            MinPercent = YearlyPriceChangePercent
            MinTickerName = TickerName
            
        End If
        
        If (TotalTickerVolume > MaxVolume) Then
            MaxVolume = TotalTickerVolume
            MaxVolmeTickerName = TickerName
        End If
        
        
        'Reset values
        YearlyPriceChangePercent = 0
        TotalTickerVolume = 0
        '
    Else
    
        TotalTickerVolume = TotalTickerVolume + MainWs.Cells(i, 7).Value
    End If
    
    Next i
    'Print Values for Greatest Increase, decrease, and volume metrics
    
    
      MainWs.Range("Q2").Value = (CStr(MaxPercent) & "%")
      MainWs.Range("Q3").Value = (CStr(MinPercent) & "%")
      MainWs.Range("P2").Value = MaxTickerName
      MainWs.Range("P3").Value = MinTickerName
      MainWs.Range("Q4").Value = MaxVolume
      MainWs.Range("O2").Value = "Greatest % Increase"
      MainWs.Range("O3").Value = "Greatest % Decrease"
      MainWs.Range("O4").Value = "Greatest Total Volume"
      
        
    
Next MainWs












End Sub
