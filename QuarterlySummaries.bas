Attribute VB_Name = "Module1"
Sub QuarterlySummaries()
    For Each Worksheet In Worksheets
        With Worksheet:
            ' Get # of rows in the data table
            Dim RowCount As Variant
            RowCount = .Cells(Rows.Count, 1).End(xlUp).Row
            
            ' Setup variables
            Dim CurrentTicker As String
            Dim OpenPrice, ClosePrice As Double
            Dim TotalStockVolume As Variant
            Dim CurrentSummaryRow As Integer
            
            Dim QuarterlyChange, PercentChange As Double
            
            ' Initialize variables
            OpenPrice = .Cells(2, 3).Value
            TotalStockVolume = 0
            CurrentSummaryRow = 2
            
            ' create summary header
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Quarterly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"

            ' Create GREATEST header
            .Range("O1").Value = "Ticker"
            .Range("P1").Value = "Value"
            .Range("N2").Value = "Greatest % Increase"
            .Range("N3").Value = "Greatest % Decrease"
            .Range("N4").Value = "Greatest Total Volume"
            
            ' Iterate through the data table
            For r = 2 To RowCount
                If .Cells(r, 1).Value = .Cells(r + 1, 1).Value Then
                    ' Same Ticker:
                    TotalStockVolume = TotalStockVolume + .Cells(r, 7)
                                       
                Else
                    ' Different Ticker:
                    ' Print summary of current Ticker
                    ClosePrice = .Cells(r, 6).Value
                    
                    CurrentTicker = .Cells(r, 1).Value
                    QuarterlyChange = ClosePrice - OpenPrice
                    PercentChange = QuarterlyChange / OpenPrice
                    TotalStockVolume = TotalStockVolume + .Cells(r, 7)
                    
                    .Cells(CurrentSummaryRow, "I").Value = CurrentTicker
                    .Cells(CurrentSummaryRow, "J").Value = QuarterlyChange
                    .Cells(CurrentSummaryRow, "K").Value = PercentChange
                    .Cells(CurrentSummaryRow, "L").Value = TotalStockVolume
                    
                    .Cells(CurrentSummaryRow, "J").NumberFormat = "0.00"
                    .Cells(CurrentSummaryRow, "K").NumberFormat = "0.00%"
                    
                    ' CONDITIONAL FORMATTING
                    ' NOTE: The following code was collected via the Record Macro feature and modified to work with the QuarterlySummaries subroutine
                    ' Quarterly Change conditional formatting
                    With .Cells(CurrentSummaryRow, "J")
                        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                            Formula1:="=0"
                        .FormatConditions(.FormatConditions.Count).SetFirstPriority
                        With .FormatConditions(1).Interior
                            .PatternColorIndex = xlAutomatic
                            .Color = 255
                            .TintAndShade = 0
                        End With
                        .FormatConditions(1).StopIfTrue = False
                        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                            Formula1:="=0"
                        .FormatConditions(.FormatConditions.Count).SetFirstPriority
                        With .FormatConditions(1).Interior
                            .PatternColorIndex = xlAutomatic
                            .Color = 5287936
                            .TintAndShade = 0
                        End With
                        .FormatConditions(1).StopIfTrue = False
                    End With
                    
                    ' Percent Change conditional formatting
                    With .Cells(CurrentSummaryRow, "K")
                        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                            Formula1:="=0"
                        .FormatConditions(.FormatConditions.Count).SetFirstPriority
                        With .FormatConditions(1).Interior
                            .PatternColorIndex = xlAutomatic
                            .Color = 255
                            .TintAndShade = 0
                        End With
                        .FormatConditions(1).StopIfTrue = False
                        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                            Formula1:="=0"
                        .FormatConditions(.FormatConditions.Count).SetFirstPriority
                        With .FormatConditions(1).Interior
                            .PatternColorIndex = xlAutomatic
                            .Color = 5287936
                            .TintAndShade = 0
                        End With
                        .FormatConditions(1).StopIfTrue = False
                    End With
                    ' END NOTE
                    
                    ' Reset for next Ticker
                    CurrentSummaryRow = CurrentSummaryRow + 1
                    OpenPrice = .Cells(r + 1, 3).Value
                    TotalStockVolume = 0
                End If
            Next r
            
            ' Find the GREATEST
            RowCount = .Cells(Rows.Count, "I").End(xlUp).Row
            
            Dim GPI_Ticker, GPD_Ticker, GTV_Ticker As String
            Dim GreatestPercentIncrease, GreatestPercentDecrease As Double
            Dim GreatestTotalVolume As Variant
            
            ' Iterate through summary table
            For r = 2 To RowCount
                If .Cells(r, "K").Value > GreatestPercentIncrease Then
                    GPI_Ticker = .Cells(r, "I").Value
                    GreatestPercentIncrease = .Cells(r, "K").Value
                ElseIf .Cells(r, "K").Value < GreatestPercentDecrease Then
                    GPD_Ticker = .Cells(r, "I").Value
                    GreatestPercentDecrease = .Cells(r, "K").Value
                End If
                
                If .Cells(r, "L").Value > GreatestTotalVolume Then
                    GTV_Ticker = .Cells(r, "I").Value
                    GreatestTotalVolume = .Cells(r, "L").Value
                End If
            Next r
            
            .Range("O2").Value = GPI_Ticker
            .Range("P2").Value = GreatestPercentIncrease
            .Range("P2").NumberFormat = "0.00%"
            
            .Range("O3").Value = GPD_Ticker
            .Range("P3").Value = GreatestPercentDecrease
            .Range("P3").NumberFormat = "0.00%"
            
            .Range("O4").Value = GTV_Ticker
            .Range("P4").Value = GreatestTotalVolume
            
            ' Autofit columns to display properly
            .Columns("J:Q").EntireColumn.AutoFit
            
            ' Reset GREATEST variables for next quarter
            GPI_Ticker = ""
            GPD_Ticker = ""
            GTV_Ticker = ""
            GreatestPercentIncrease = 0
            GreatestPercentDecrease = 0
            GreatestTotalVolume = 0
        End With
    Next Worksheet
End Sub
