# VBA-challenge
HW 2: Use VBA scripting to analyze real stock market data
Sub WallStreet()
    'stating variables
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        Dim ticker As String
        Dim startmoney As Double
        Dim endmoney As Double
        Dim change As Double
        Dim percent As Double
        Dim stockvolume As Double
        Dim i As Long
        Dim j As Long
        stockvolume = Cells(2, 7).Value
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        startmoney = Cells(2, 3).Value
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        j = 2
        ticker = Cells(2, 1).Value
        For i = 3 To lastrow
            If Cells(i, 1).Value = Cells(i - 1, 1).Value Then
                If startmoney = 0 Then
                    startmoney = Cells(i, 6).Value
                End If
                stockvolume = stockvolume + Cells(i, 7).Value
                ticker = Cells(i, 1).Value
            Else
                Debug.Print ticker
                Cells(j, 12).Value = stockvolume
                endmoney = Cells(i - 1, 6).Value
                change = endmoney - startmoney
                If startmoney = 0 Then
                    percent = 0
                Else
                    percent = change / startmoney
                End If
                stockvolume = Cells(i, 7).Value
                startmoney = Cells(i, 3).Value
                Cells(j, 9).Value = ticker
                Cells(j, 10).Value = change
                    If (change < -1) Then
                        Cells(j, 10).Interior.ColorIndex = 3
                    Else
                        Cells(j, 10).Interior.ColorIndex = 4
                    End If
                Cells(j, 11) = percent
                Cells(j, 11).NumberFormat = "0.00%"
                percent = 0
                change = 0
                j = j + 1
                ticker = Cells(i - 1, 1).Value
            End If
        Next i
    Next
End Sub
