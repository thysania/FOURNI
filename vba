Sub GenerateReceipts()
    Dim wsData As Worksheet, wsReceipt As Worksheet
    Dim lastRow As Long, i As Long
    Dim receiptRow As Long
    Dim clientName As String, unitID As String
    Dim address As String, amount As Currency, dueDate As String
    
    Set wsData = ThisWorkbook.Sheets("Monthly Rent")
    Set wsReceipt = ThisWorkbook.Sheets("Receipt")
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRow ' Assuming headers in row 1
        If wsData.Cells(i, "F").Value <> "Y" Then ' Unpaid only
            ' Get data
            clientName = wsData.Cells(i, "A").Value
            unitID = wsData.Cells(i, "B").Value
            address = wsData.Cells(i, "C").Value
            amount = wsData.Cells(i, "D").Value
            dueDate = wsData.Cells(i, "E").Value
            
            ' Fill the receipt template
            wsReceipt.Range("ClientName").Value = clientName
            wsReceipt.Range("Unit").Value = unitID
            wsReceipt.Range("Address").Value = address
            wsReceipt.Range("Amount").Value = amount
            wsReceipt.Range("Due").Value = dueDate
            
            ' Optional: Print or Save PDF
            ' wsReceipt.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            '     ThisWorkbook.Path & "\Receipt_" & clientName & "_" & unitID & ".pdf"
            ' wsReceipt.PrintOut
            
            ' Optional: mark as paid
            ' wsData.Cells(i, "F").Value = "Y"
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Receipts generated!"
End Sub