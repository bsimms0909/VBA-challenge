Attribute VB_Name = "Module1"
Sub stock_analysis():
    Dim Counter As Integer
    Dim lastrow As Long
    Dim volume As Double
    Dim ticker As String
    Dim Next_Ticker As String
    Dim Previous_Ticker As String
    Dim output_row As Integer
    Dim Year As Integer
    Dim Closing_Price As Double
    Dim Opening_Price As Double
    
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Increase_Ticker As String
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Percent_Decrease_Ticker As String
    Dim Greatest_Volume As String
    Dim Greatest_Volume_Ticker As String
    Dim Percent_Change As Double
    
    For Year = 2018 To 2020
    
        Sheets(CStr(Year)).Select
        
        last_row = Cells(Rows.Count, "A").End(xlUp).Row
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        volume = 0
        output_row = 2
        
        Greatest_Percent_Increase = 0
        Greatest_Percent_Decrease = 0
        Greatest_Volume = 0
        
        For i = 2 To last_row
            volume = volume + Cells(i, 7).Value
            ticker = Cells(i, 1).Value
            Next_Ticker = Cells(i + 1, 1).Value
            Previous_Ticker = Cells(i - 1, 1).Value
            If ticker <> Next_Ticker Then
            Closing_Price = Cells(i, 6).Value
                Range("I" & output_row).Value = ticker
                Range("L" & output_row).Value = volume
                Range("J" & output_row).Value = Closing_Price - Opening_Price
                If Closing_Price - Opening_Price >= 0 Then
                    Range("J" & output_row).Interior.Color = RGB(0, 255, 0)
                Else
                    Range("J" & output_row).Interior.Color = RGB(255, 0, 0)
                    End If
                Range("K" & output_row).Value = (Closing_Price - Opening_Price) / Opening_Price
                volume = 0
                output_row = output_row + 1
                ElseIf ticker <> Previous_Ticker Then
            Opening_Price = Cells(i, 3).Value
        End If
            
        Next i
        'Row_Count = Cells(Rows.Count, "J").End(xlUp).Row
        Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & last_row)) * 100
        Increase_Number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & last_row)), Range("K2:K" & last_row), 0)
        Range("P2").Value = Cells(Increase_Number + 1, 9)
        
        Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & last_row)) * 100
        Decrease_Number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & last_row)), Range("K2:K" & last_row), 0)
        Range("P3").Value = Cells(Decrease_Number + 1, 9)
        
        Range("Q4") = WorksheetFunction.Max(Range("L2:L" & last_row))
        volume = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & last_row)), Range("L2:L" & last_row), 0)
        Range("P4").Value = Cells(volume + 1, 9)
        
        Next Year
        
    'last_row = Cells(Rows.Count, "K").End(xlIp).Row
    MsgBox ("completed")
    

End Sub

   
