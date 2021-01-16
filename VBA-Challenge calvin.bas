Attribute VB_Name = "Module1"

Sub Stock_Analysis()
    
 
    Dim ws As Worksheet

    For Each ws In Worksheets

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        Dim ticker As String
        Dim total_vol As Double
        Dim rowcount As Long
        Dim openprice As Double
        Dim closeprice As Double
        Dim year_change As Double
        Dim percent_change As Double
        
        
        
        total_vol = 0
        rowcount = 2
        openprice = 0
        closeprice = 0
        year_change = 0
        percent_change = 0

        
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        For i = 2 To lastrow
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                openprice = ws.Cells(i, 3).Value

            End If

            total_vol = total_vol + ws.Cells(i, 7)

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(rowcount, 12).Value = total_vol

            
                closeprice = ws.Cells(i, 6).Value

                
                year_change = closeprice - openprice
                ws.Cells(rowcount, 10).Value = year_change

                

            
                If openprice = 0 And closeprice = 0 Then
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                   
                ElseIf openprice = 0 Then
                    Dim openzero As String
                    openzero = " "
                    ws.Cells(rowcount, 11).Value = openzero
                Else
                    percent_change = year_change / openprice
                    ws.Cells(rowcount, 11).Value = percent_change
                    
                End If

        
                rowcount = rowcount + 1

                total_vol = 0
                openprice = 0
                closeprice = 0
                year_change = 0
                percent_change = 0
                
                
                If year_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If
                 ws.Cells(rowcount, 11).NumberFormat = "0.00%"
            End If
        Next i
    Next ws
End Sub

