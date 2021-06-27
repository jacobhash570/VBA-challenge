Attribute VB_Name = "Module1"
Sub Stocks()


  Dim ticker As String

  Dim vol As Double
  vol = 0


  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  LastRow = Cells(Rows.Count, "A").End(xlUp).Row

  For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ticker = Cells(i, 1).Value

      vol = vol + Cells(i, 7).Value
     
      Range("K" & Summary_Table_Row).Value = ticker

       Range("L" & Summary_Table_Row).Value = vol

      Summary_Table_Row = Summary_Table_Row + 1

       vol = 0

    Else

      vol = vol + Cells(i, 7).Value

    End If

  Next i

End Sub
