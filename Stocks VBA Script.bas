Attribute VB_Name = "Module1"
Sub Stocks():

Dim current As Worksheet

For Each current In ThisWorkbook.Worksheets

current.Activate

Dim row As Integer
Dim count As LongLong
Dim openprice As Double
Dim rowcount As Long

openprice = 0
count = 0
row = 2

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Columns("J").ColumnWidth = 12
Columns("K:L").ColumnWidth = 20
Columns("O").ColumnWidth = 20
Columns("Q").ColumnWidth = 12

rowcount = Range("A1").End(xlDown).row

For I = 2 To rowcount

    count = count + Cells(I, 7).Value

    If Cells(I, 1).Value = Cells(I + 1, 1) Then

       If openprice = 0 Then

            openprice = Cells(I, 3).Value

        End If

    Else:

        Cells(row, 9).Value = Cells(I, 1).Value

        Cells(row, 10).Value = Cells(I, 6).Value - openprice

        If Cells(row, 10).Value < 0 Then

            Cells(row, 10).Interior.ColorIndex = 3

        Else:

            Cells(row, 10).Interior.ColorIndex = 4

        End If

        Cells(row, 11).Value = FormatPercent(Cells(row, 10).Value / openprice, 2)

        Cells(row, 12).Value = count

        row = row + 1

        count = 0

        openprice = 0

    End If

Next I

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

rowcount = Range("I1").End(xlDown).row

Cells(2, 17).Value = -999
Cells(3, 17).Value = 999
Cells(4, 17).Value = 0

For I = 2 To rowcount

    If Cells(I, 11).Value > Cells(2, 17).Value Then

        Cells(2, 17).Value = FormatPercent(Cells(I, 11).Value, 2)

        Cells(2, 16).Value = Cells(I, 9).Value

    End If

    If Cells(I, 11).Value < Cells(3, 17).Value Then

        Cells(3, 17).Value = FormatPercent(Cells(I, 11).Value, 2)

        Cells(3, 16).Value = Cells(I, 9).Value

    End If

    If Cells(I, 12).Value > Cells(4, 17).Value Then

        Cells(4, 17).Value = Cells(I, 12).Value

        Cells(4, 16).Value = Cells(I, 9).Value

    End If

Next I

'MsgBox (current.Name)'

Next current

End Sub

