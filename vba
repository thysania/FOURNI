Sub PrintReceiptsFromColumns()
    Dim wsReceipt As Worksheet, wsData As Worksheet
    Dim lastRow As Long, i As Long

    Set wsReceipt = ThisWorkbook.Sheets("RECPT")
    Set wsData = ThisWorkbook.Sheets("Monthly Rent") ' Adjust if your data sheet is named differently

    ' Assumes data starts from row 2 (row 1 is headers)
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        ' Write current row's values into the fixed positions
        wsReceipt.Range("B22").Value = wsData.Cells(i, 2).Value  ' MOIS (column B)
        wsReceipt.Range("B6").Value = Date                       ' Today()
        wsReceipt.Range("F7").Value = wsData.Cells(i, 3).Value   ' MT (column C)
        wsReceipt.Range("D14").Value = wsData.Cells(i, 1).Value  ' LOCATAIRE (column A)

        ' Print the filled receipt
        wsReceipt.PrintOut ' Or use ExportAsFixedFormat to save as PDF
    Next i

    Application.ScreenUpdating = True
    MsgBox "All receipts printed!"
End Sub