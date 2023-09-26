Sub VBAChallenge()
Dim i As Long
Dim ws_num As Integer
Dim ws As Worksheet
ws_num = ThisWorkbook.Worksheets.Count


Application.ScreenUpdating = False


For w = 1 To ws_num

Sheets(w).Activate

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("I2:AA10000").Clear

  Dim Ticker As String
  Dim Dat As String
  Dim Open_price, Close_price, Tot_volume As Double
  Dim Sum_Row As Integer
  Dim current As String
  


    nrows = Range("A2").End(xlDown).Row
    Ticker = Cells(2, 1).Value
    current = Cells(2, 1).Value
    Open_price = Cells(2, 3).Value
    Dat = Cells(2, 2).Value
    Tot_volume = 0
    Sum_Row = 2
    
  For i = 2 To nrows
   
    If Cells(i + 1, 1).Value <> current Then
 
      
      Close_price = Cells(i, 6).Value
      Tot_volume = Tot_volume + Cells(i, 7).Value
      
      'Calculations
      
      Range("I" & Sum_Row).Value = Ticker
      Range("J" & Sum_Row).Value = Close_price - Open_price
      Range("K" & Sum_Row).Value = (Close_price - Open_price) / Open_price
      Range("L" & Sum_Row).Value = Tot_volume
      
      Sum_Row = Sum_Row + 1
      ' Reset the variables
      Ticker = Cells(i + 1, 1).Value
      Close_price = 0
      current = Cells(i + 1, 1).Value
      Open_price = Cells(i + 1, 3).Value
      Tot_volume = 0
 
    Else
      ' Add to the Tot_volume
      Tot_volume = Tot_volume + Cells(i, 7).Value
    End If
  Next i
  

'Call the calculation of percentage
Call percentagesum


Next

'Call macro to put format
Call format

End Sub


Sub percentagesum()

  Dim Ticker As String
  Dim Percentage, Volume As Double
  Dim Sum_Row As Integer
  Dim current As String
  
  Range("O1").Value = "Ticker"
  Range("P1").Value = "Value"
  Range("N2").Value = "Greatest % increase"
  Range("N3").Value = "Greatest % decrease"
  Range("N4").Value = "Greatest Total Volume"
  
  

    nrows = Range("I2").End(xlDown).Row
    
        Ticker_max = Cells(2, 9).Value
        Ticker_min = Cells(2, 9).Value
        porcentage_max = Cells(2, 11).Value
        porcentage_min = Cells(2, 11).Value
        Ticker_rev = Cells(2, 9).Value
        Tot_volume_rev = Cells(2, 12).Value

For i = 2 To nrows

        
    
    If Cells(i + 1, 11).Value > porcentage_max Then
        Ticker_max = Cells(i + 1, 9).Value
        porcentage_max = Cells(i + 1, 11).Value
        Range("O2").Value = Ticker_max
        Range("P2").Value = porcentage_max
    
        
        
    Else
        Range("O3").Value = Ticker_max
        Range("P3").Value = porcentage_max

    End If
    
    
    
    
    If Cells(i + 1, 11).Value < porcentage_min Then
        Ticker_min = Cells(i + 1, 9).Value
        porcentage_min = Cells(i + 1, 11).Value
        Range("O3").Value = Ticker_min
        Range("P3").Value = porcentage_min
    Else
        Range("O3").Value = Ticker_min
        Range("P3").Value = porcentage_min
    
    End If
    
    
    If Cells(i + 1, 12).Value > Tot_volume_rev Then
        Ticker_rev = Cells(i + 1, 9).Value
        Tot_volume_rev = Cells(i + 1, 12).Value
        Range("O4").Value = Ticker_rev
         Range("P4").Value = Tot_volume_rev
         
    Else
        Range("O4").Value = Ticker_rev
        Range("P4").Value = Tot_volume_rev
        
    End If
    


Next i

  Range("P2:P3").Style = "Percent"
  Range("P4").Style = "Comma"


End Sub

Sub format()


Dim i As Long
Dim nrows As Long
Dim ws_num As Integer
Dim ws As Worksheet
ws_num = ThisWorkbook.Worksheets.Count

For w = 1 To ws_num

Sheets(w).Activate


nrows = Range("I2").End(xlDown).Row

For i = 2 To nrows
        Value = Range("J" & i).Value

    'Add the format in the cell
      If Value <= 0 Then
        
        Cells(i, 10).Interior.Color = RGB(255, 0, 0)
     Else
        Range("J" & i).Interior.ColorIndex = 43
     End If

Next i


  Columns("K:K").Style = "Percent"
  Columns("L:L").Style = "Comma"
Next
  Sheets(1).Activate
  Range("J2").Activate
  
  
  MsgBox ("Finished!!")

End Sub


