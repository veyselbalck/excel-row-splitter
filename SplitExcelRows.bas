Sub SplitExcelRows()
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim rowsPerSheet As Long

    rowsPerSheet = 5000
    Set sourceSheet = ActiveSheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    startRow = 1

    Do While startRow <= lastRow
        Set newSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        newSheet.Name = "Sheet " & Sheets.Count - 1

        endRow = startRow + rowsPerSheet - 1
        If endRow > lastRow Then
            endRow = lastRow
        End If

        sourceSheet.Rows(startRow & ":" & endRow).Copy newSheet.Rows(1)
        startRow = endRow + 1
    Loop
End Sub
