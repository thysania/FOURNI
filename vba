Sub PreviewOneReceipt()
    Dim wsReceipt As Worksheet, wsData As Worksheet
    Dim selectedRow As Variant

    Set wsReceipt = ThisWorkbook.Sheets("RECPT")
    Set wsData = ThisWorkbook.Sheets("RENT")

    selectedRow = Application.InputBox("Enter row number to preview:", Type:=1)
    If selectedRow = False Then Exit Sub
    If selectedRow < 2 Or selectedRow > wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row Then
        MsgBox "Invalid row selected."
        Exit Sub
    End If

    ' Fill in receipt
    wsReceipt.Range("B22").Value = wsData.Cells(selectedRow, 2).Value ' MOIS
    wsReceipt.Range("B6").Value = Date
    wsReceipt.Range("F7").Value = wsData.Cells(selectedRow, 3).Value ' MT
    wsReceipt.Range("D14").Value = wsData.Cells(selectedRow, 1).Value ' LOCATAIRE

    ' Show Excel's built-in print preview for this receipt
    wsReceipt.PrintPreview
End Sub