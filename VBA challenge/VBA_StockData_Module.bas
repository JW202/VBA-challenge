Attribute VB_Name = "Module1"
Sub StockData()

Dim I, J As Long
Dim ticker As String
Dim Openprice, Closeprice, Volume As Double



ticker = Cells(2, 1)
Openprice = Cells(2, 3)
Volume = Cells(2, 7)


J = 2
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"
Cells(2, 9) = ticker

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 3 To LastRow
    If Cells(I, 1) <> 0 Then
        If Cells(I, 1) <> ticker Then
            Closeprice = Cells(I - 1, 6)
            Cells(J, 10) = Closeprice - Openprice
            If Cells(J, 10) >= 0 Then
                   Cells(J, 10).Interior.ColorIndex = 4
            Else
                    Cells(J, 10).Interior.ColorIndex = 3
            
            
            End If
            
            Cells(J, 11) = Cells(J, 10) / Openprice
            Cells(J, 12) = Volume
            ticker = Cells(I, 1)
            Openprice = Cells(I, 3)
            Volume = Cells(I, 7)
            J = J + 1
            Cells(J, 9) = ticker
            
        Else
            Volume = Volume + Cells(I, 7)
        End If
    
            
    End If


Next I

   Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    
    Columns("j:j").Select
    Selection.NumberFormat = "0.00"

' Bonus question

   Dim Increase, Decrease, GVolume As Double
   Dim Iticker, Dticker, Vticker As String
   Dim K As Long
   
   Increase = Cells(2, 11)
   Decrease = Cells(2, 11)
   GVolume = Cells(2, 12)
   Iticker = Cells(2, 9)
   Dticker = Cells(2, 9)
   Vticker = Cells(2, 9)
   
   For K = 3 To J
   
     If Cells(K, 11) > Increase Then
        Increase = Cells(K, 11)
        Iticker = Cells(K, 9)
     End If
     If Cells(K, 11) < Decrease Then
        Decrease = Cells(K, 11)
        Dticker = Cells(K, 9)
    End If
    If Cells(K, 12) > GVolume Then
      GVolume = Cells(K, 12)
      Vticker = Cells(K, 9)
    End If
   Next K
   
   Cells(2, 16) = Iticker
   Cells(2, 17) = Increase
   Cells(3, 16) = Dticker
   Cells(3, 17) = Decrease
   Cells(4, 16) = Vticker
   Cells(4, 17) = GVolume
   
    
End Sub
