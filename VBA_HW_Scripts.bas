Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()
    
    'Print column headers
    Range("I1").Value = "Ticker Symbol"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    
    ' Variable Declaratoins
    Dim ticker As String
    Dim volume As LongLong
    Dim lastrow As Long
    Dim opn As Double
    Dim cls As Double
    Dim tcount As Integer
    Dim pct_chg As Double
    Dim ann_chg As Double
    
    
    'Variable Assignments
    ticker = Cells(2, 1).Value
    volume = Cells(2, 7).Value
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    opn = Cells(2, 3).Value
    cls = 0
    tcount = 1

  
   
    'For loop with conditional through each ticker, capturing the requested information and printing in the specefied locations [I,J,K,L].
    For i = 2 To lastrow
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            cls = Cells(i, 6).Value
            tcount = tcount + 1
            volume = volume + Cells(i, 7).Value
            ann_chg = (cls - opn)
            pct_chg = (cls - opn) / cls
           
        
            'Print
            Cells(tcount, 9).Value = ticker
            Cells(tcount, 12).Value = volume
                If ann_chg > 0 Then
                    Cells(tcount, 10).Value = ann_chg
                    Cells(tcount, 10).Interior.ColorIndex = 4
                ElseIf ann_chg < 0 Then
                    Cells(tcount, 10).Value = ann_chg
                    Cells(tcount, 10).Interior.ColorIndex = 3
                End If
            Cells(tcount, 11).Value = Round(pct_chg, 2)
            Cells(tcount, 11).NumberFormat = "0.00%"
            

            ' --Set the next ticker
            ticker = Cells(i + 1, 1).Value
            volume = 0
            opn = Cells(i + 1, 3).Value
            
        Else
            'Add to the volume total
            volume = volume + Cells(i, 7).Value
        
        End If
    Next i

End Sub
    
     
