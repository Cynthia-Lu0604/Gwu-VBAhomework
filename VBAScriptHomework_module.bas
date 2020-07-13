Attribute VB_Name = "Module1"
Sub stocktrend():


Dim ticker As String
Dim numberticker As Integer
Dim lastrow As Long
Dim openingprice As Double
Dim closingprice As Double
Dim yrchange As Double
Dim perchange As Double
Dim totalstockvolume As Double

'loopthru workbook
For Each ws In Worksheets

    
    ws.Activate

    ' find last row
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add header
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' reset variables
    numberticker = 0
    ticker = ""
    yerchange = 0
    openingprice = 0
    perchange = 0
    totalstockvolume = 0
    
    ' loop begins
    For i = 2 To lastrow

        ' retreive ticker
        ticker = Cells(i, 1).Value
        
        ' retreive opening price
        If openingprice = 0 Then
            openingprice = Cells(i, 3).Value
        End If
        
        ' Add up the total stock volume
        totalstockvolume = totalstockvolume + Cells(i, 7).Value
        
        
        If Cells(i + 1, 1).Value <> ticker Then
    
            numberticker = numberticker + 1
            Cells(numberticker + 1, 9) = ticker
            
            ' retreive closingprice
            closingprice = Cells(i, 6)
            
            ' calculate yearly change
            yrchange = closingprice - openingprice
            
            ' insert yearly change
            Cells(numberticker + 1, 10).Value = yrchange
            
            ' colorize
            If yrchange > 0 Then
                Cells(numberticker + 1, 10).Interior.ColorIndex = 4
        
            ElseIf yrchange < 0 Then
                Cells(numberticker + 1, 10).Interior.ColorIndex = 3
           
            Else
                Cells(numberticker + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Calculate percent change
            If openingprice = 0 Then
                perchange = 0
            Else
                perchange = (yrchange / openingprice)
            End If
            
            
            ' display in percentage form
            Cells(numberticker + 1, 11).Value = Format(perchange, "Percent")
            
            
            ' reset opening price
            openingprice = 0
            
            ' retreive total stock volume
            Cells(numberticker + 1, 12).Value = totalstockvolume
            
            ' Set total stock volume to 0
            totalstockvolume = 0
        End If
        
    Next i
    
    
Next ws


End Sub
