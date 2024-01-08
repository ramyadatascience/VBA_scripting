VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub RunMacroOnAllSheetsToRight()

For i = ActiveSheet.Index To Sheets.Count
    Stockanalysis i
Next i

End Sub


Sub Stockanalysis(ByVal Sheets_Index As Integer)

Dim Ws As Worksheet
Set Ws = Worksheets(Sheets_Index)


Dim Counterstart As Long
Dim Counterend  As Long
Dim Lastrow        As Long
Dim Yearchange  As Double
Dim Counter       As Long
Dim Stockvol      As Long


Lastrow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
Counter = 1


   Counterstart = 2
   
   Ws.Range("I1").EntireColumn.Insert
   Ws.Cells(1, 9).Value = "Ticker"
    
  Ws.Range("J1").EntireColumn.Insert
   Ws.Cells(1, 10).Value = "YearlyChange"
    
  Ws.Range("K1").EntireColumn.Insert
   Ws.Cells(1, 11).Value = "Percent Change"
   
  Ws.Range("L1").EntireColumn.Insert
   Ws.Cells(1, 12).Value = "Total Stock Volume"
   
   
   For i = 2 To Lastrow
   
    If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
    
        Counterend = i
        Counter = Counter + 1
        Yearchange = Ws.Cells(Counterend, 3).Value - Ws.Cells(Counterstart, 3).Value
        Ws.Cells(Counter, 9).Value = Ws.Cells(i, 1).Value
        Ws.Cells(Counter, 10).Value = Yearchange
        Ws.Cells(Counter, 10).Interior.ColorIndex = 4
        
        If Yearchange < 0 Then
           
           Ws.Cells(Counter, 10).Interior.ColorIndex = 3
           
        End If
        ' MsgBox (Yearchange)
        Ws.Cells(Counter, 11).Value = Yearchange / Ws.Cells(Counterstart, 3).Value
        Ws.Cells(Counter, 11).NumberFormat = "00.0%"
        Ws.Cells(Counter, 12).Value = WorksheetFunction.Sum(Ws.Range(Ws.Cells(Counterstart, 7), Ws.Cells(Counterend, 7)))
        Ws.Cells(Counter, 12).NumberFormat = "00"
        ' MsgBox ("sum")
    End If
  
   Next i
   
   
  Ws.Range("O1").EntireColumn.Insert
   Ws.Cells(1, 15).Value = "Ticker"
    
  Ws.Range("P1").EntireColumn.Insert
   Ws.Cells(1, 16).Value = "Value"
   
   Ws.Cells(2, 14).Value = "Greatest % Increase"
   Ws.Cells(3, 14).Value = "Greatest % Decrease"
   Ws.Cells(4, 14).Value = "Greatest Total Volume"
       
       Range1 = Ws.Range(Ws.Cells(2, 11), Ws.Cells(Counter, 11))
       Range2 = Ws.Range(Ws.Cells(2, 12), Ws.Cells(Counter, 12))
       
       'MsgBox (Counter)
       
       Dim MaxValue As Double
       Dim MinValue As Double
       Dim Maxstock As Double
       
       MaxValue = WorksheetFunction.Max(Range1)
       MinValue = WorksheetFunction.Min(Range1)
       Maxstock = WorksheetFunction.Max(Range2)
       
       
       Ws.Cells(2, 16).Value = MaxValue
       Ws.Cells(2, 16).NumberFormat = "00.0%"
       
       Ws.Cells(3, 16).Value = MinValue
       Ws.Cells(3, 16).NumberFormat = "00.0%"
       
       Ws.Cells(4, 16).Value = Maxstock
       Ws.Cells(4, 16).NumberFormat = "00"

 For i = 2 To Counter

    If Ws.Cells(i, 11).Value = MaxValue Then
   
        Ws.Cells(2, 15).Value = Ws.Cells(i, 9).Value
    
    End If
            
    If Ws.Cells(i, 11).Value = MinValue Then
   
        Ws.Cells(3, 15).Value = Ws.Cells(i, 9).Value
    
    End If
     
        If Ws.Cells(i, 12).Value = Maxstock Then
   
        Ws.Cells(4, 15).Value = Ws.Cells(i, 9).Value
    
    End If
     
     Next i



End Sub

