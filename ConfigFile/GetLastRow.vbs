Function GetLastRow(sheetName As String, colNumber As Integer) As Long
    GetLastRow = Sheets(sheetName).Cells(Rows.Count, colNumber).End(xlUp).Row
End Function