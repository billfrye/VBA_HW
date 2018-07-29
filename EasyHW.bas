Attribute VB_Name = "EasyHW"

'Easy version
'* Create a script that will loop through each year of stock data and
'  grab the total amount of volume each stock had over the year.
'* You will also need to display the ticker symbol to coincide with the total volume.
'* Your result should look as follows (note: all solution images are for 2015 data).
'  ![easy_solution](Images/easy_solution.png)

Sub DoEasyHw()
    
    Dim In_TickerCol, In_VolumeCol As Integer
    
    Dim Out_TickerCol, Out_TickerVolumeCol As Integer
    Dim Out_Ticker As String
    Dim Out_TickerVolumeSum As Double
    
    In_TickerCol = 1
    In_VolumeCol = 7
    Out_TickerCol = 9
    Out_TickerVolumeCol = 10
    
    Dim ws As Worksheet
    Dim row, Out_Row As Long
    Dim lastrow As Long
    
    For Each ws In Sheets
        ' Start of a new worksheet...
        If (Left(ws.Name, 4) = "Test") Then
            'Skip this sheet...
        Else
            lastrow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
            ws.Cells(1, Out_TickerCol).Value = "Ticket"
            ws.Cells(1, Out_TickerVolumeCol).Value = "Ticker Stock Volume"
            
            Out_Row = 2
            ws.Cells(Out_Row, Out_TickerCol).Value = Out_Ticker
            'ws.Cells(2, In_TickerCol).Value = Out_Ticker
            Out_TickerVolumeSum = 0
            For row = 2 To lastrow
                If Out_Ticker <> ws.Cells(row, In_TickerCol).Value Then
                    'starting a new ticker, save the current sum & ticker
                    'and start summing again
                    ws.Cells(Out_Row, Out_TickerCol).Value = Out_Ticker
                    ws.Cells(Out_Row, Out_TickerVolumeCol).Value = Out_TickerVolumeSum
                    'Setup the new ticker value & volume sum
                    Out_Ticker = ws.Cells(row, In_TickerCol).Value
                    Out_TickerVolumeSum = 0
                    Out_Row = Out_Row + 1
                End If
                Out_TickerVolumeSum = Out_TickerVolumeSum + ws.Cells(row, In_VolumeCol).Value
                
            Next row
            'document the last ticker & volume for this worksheet
            ws.Cells(Out_Row, Out_TickerCol).Value = Out_Ticker
            ws.Cells(Out_Row, Out_TickerVolumeCol).Value = Out_TickerVolumeSum
            'Out_TickerVolumeSum = 0
             
        End If
            
        
    Next ws
    
    
End Sub
