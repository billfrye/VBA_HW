Attribute VB_Name = "EasyHW"

'Easy version
'* Create a script that will loop through each year of stock data and
'  grab the total amount of volume each stock had over the year.
'* You will also need to display the ticker symbol to coincide with the total volume.
'* Your result should look as follows (note: all solution images are for 2015 data).
'  ![easy_solution](Images/easy_solution.png)

Option Explicit

Enum InCol
    TickerID = 1
    TransDate = 2
    DayOpen = 3
    DayClose = 6
    Volume = 7
End Enum

Enum OutCol
    TickerID = 9
    Volume = 10
End Enum

    
Sub DoEasyHW()

    Dim Out_Ticker As String
    Dim Out_TickerVolumeSum As Double
    
    Dim ws As Worksheet
    Dim InRow, OutRow As Long
    Dim LastRow As Long
    
    OutRow = 2
    For Each ws In Sheets
        ' Start of a new worksheet...
        If (Left(ws.Name, 4) = "Test") Then
            'Skip this sheet...
        Else
            LastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
            ws.Cells(1, OutCol.TickerID).Value = "Ticket"
            ws.Cells(1, OutCol.Volume).Value = "Ticker Stock Volume"
            
            'ws.Cells(OutRow, OutCol.TickerID).Value = Out_Ticker
            'ws.Cells(2, InCol.Ticker).Value = Out_Ticker
            Out_TickerVolumeSum = 0
            For InRow = 2 To LastRow
                If Out_Ticker <> ws.Cells(InRow, InCol.TickerID).Value Then
                    'starting a new ticker, save the current sum & ticker
                    'and start summing again
                    'the first time, Out_Ticker will be "", so skip writing
                    'out a row for that...
                    ' Complete test ot OutTicker = "" only needs to be done
                    ' when OutRow = 2
                    If (Out_Ticker = "") Then
                        Out_Ticker = ws.Cells(InRow, InCol.TickerID).Value
                    Else
                        ws.Cells(OutRow, OutCol.TickerID).Value = Out_Ticker
                        ws.Cells(OutRow, OutCol.Volume).Value = Out_TickerVolumeSum
                        ' Set up for the next ticker...
                        Out_Ticker = ws.Cells(InRow, InCol.TickerID).Value
                        OutRow = OutRow + 1
                    End If
                    'Setup the new ticker value & volume sum
                    Out_TickerVolumeSum = 0
                End If
                Out_TickerVolumeSum = Out_TickerVolumeSum + ws.Cells(InRow, InCol.Volume).Value
                
            Next InRow
             
        End If
                
       'document the last ticker & volume for this worksheet
        ws.Cells(OutRow, OutCol.TickerID).Value = Out_Ticker
        ws.Cells(OutRow, OutCol.Volume).Value = Out_TickerVolumeSum
       'Out_TickerVolumeSum = 0
        Out_Ticker = ""
        OutRow = 2
    Next ws
    
End Sub
