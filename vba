 Sub SumUntilBoldAndMakeFormula()
    Dim r As Long
    Dim lastRow As Long
    Dim sum As Double
    Dim col As Variant
    col = Array("T", "U", "V")
    For Each c In col
    'Find the last row with data in column c
    lastRow = Cells(Rows.Count, c).End(xlUp).Row

    'Initialize the sum to 0
    sum = 0

    'Start at row 2
    For r = 2 To lastRow + 1
        'Check if the cell in column c is empty
        If IsEmpty(Cells(r, c)) Then
            'Sum the values above the current row until a bold cell is encountered
            For i = r - 1 To 2 Step -1
                If Cells(i, c).Font.Bold Then
                    Exit For
                End If
                sum = sum + Cells(i, c).Value
            Next i
            'Insert the sum into the current empty row in column c
            With Cells(r, c)
                .Value = sum
                .Font.Bold = True
            End With
            'Add the formula in column AI
            Cells(r, "AI").Formula = "=(" & "T" & r & "+" & "U" & r & "+" & "V" & r & ")*1000"
            'Add the formula in column AJ
            Cells(r, "AJ").Formula = "=(AI" & r & "/33)*12"
            'Add the formula in column AK
            Cells(r, "AK").Formula = "=(AJ" & r & "*5)/100"
            'Add the formula in column AL
            Cells(r, "AL").Formula = "=(AJ" & r & "*10)/100"
            'Add the formula in column AM
            Cells(r, "AM").Formula = "=(AJ" & r & "*15)/100"
            'Add the formula in column AN
            Cells(r, "AN").Formula = "=(AJ" & r & "*20)/100"
            'Add the formula in column AO
            Cells(r, "AO").Formula = "=(AJ" & r & "*25)/100"
            'Add the formula in column AP
            Cells(r, "AP").Formula = "=(AJ" & r & "*30)/100"
            'Add the formula in column AQ
            Cells(r, "AQ").Formula = "=(AJ" & r & "*30)/100"
            'Color the entire row in which it entered data gray
            Rows(r).Interior.Color = RGB(192, 192, 192)
            sum = 0
        End If
    Next r
    Next c
End Sub
