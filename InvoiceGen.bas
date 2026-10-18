Attribute VB_Name = "InvoiceGen"
' Invoice Generator for project teams — by Muhammad Bilal
' Reads client + line-item data from the "Invoice-Data" sheet, then creates
' one formatted invoice sheet per client (invoice numbers INV-2026-001...).
' Re-running deletes previously generated invoice sheets first (idempotent).
' Written with plain Range/Cells operations so it runs in Excel and in
' LibreOffice (Option VBASupport 1).

Sub GenerateInvoices()
    Dim wb As Object, wsData As Object, wsInv As Object
    Dim clientRow As Long, itemRow As Long
    Dim lastClientRow As Long, lastItemRow As Long
    Dim invNum As Integer, itemCount As Integer, r As Long
    Dim clientID As String, clientName As String
    Dim clientAddr As String, clientCity As String
    Dim taxPct As Double, subtotal As Double, taxAmt As Double
    Dim grandTotal As Double, lineTotal As Double, qty As Double, rate As Double
    Dim i As Long

    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("Invoice-Data")

    ' --- remove previously generated invoice sheets (safe re-run) ---
    Application.DisplayAlerts = False
    For i = wb.Sheets.Count To 1 Step -1
        If Left(wb.Sheets(i).Name, 4) = "INV-" Then
            wb.Sheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True

    ' --- data bounds: clients in A:E, items in G:J ---
    lastClientRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastItemRow = wsData.Cells(wsData.Rows.Count, 7).End(xlUp).Row

    invNum = 0
    For clientRow = 3 To lastClientRow
        clientID = Trim(CStr(wsData.Cells(clientRow, 1).Value))
        If clientID <> "" Then
            clientName = CStr(wsData.Cells(clientRow, 2).Value)
            clientAddr = CStr(wsData.Cells(clientRow, 3).Value)
            clientCity = CStr(wsData.Cells(clientRow, 4).Value)
            taxPct = CDbl(wsData.Cells(clientRow, 5).Value)
            invNum = invNum + 1

            Set wsInv = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            wsInv.Name = "INV-2026-" & Format(invNum, "000")
            Call PaintInvoice(wsInv)

            ' --- header block ---
            wsInv.Range("C3").Value = wsInv.Name
            wsInv.Range("C4").Value = Date
            wsInv.Range("C4").NumberFormat = "yyyy-mm-dd"
            wsInv.Range("C5").Value = clientName
            wsInv.Range("C6").Value = clientAddr & ", " & clientCity

            ' --- line items ---
            r = 9
            itemCount = 0
            subtotal = 0
            For itemRow = 3 To lastItemRow
                If Trim(CStr(wsData.Cells(itemRow, 7).Value)) = clientID Then
                    itemCount = itemCount + 1
                    qty = CDbl(wsData.Cells(itemRow, 9).Value)
                    rate = CDbl(wsData.Cells(itemRow, 10).Value)
                    lineTotal = qty * rate
                    wsInv.Cells(r, 1).Value = itemCount
                    wsInv.Cells(r, 2).Value = CStr(wsData.Cells(itemRow, 8).Value)
                    wsInv.Cells(r, 3).Value = qty
                    wsInv.Cells(r, 4).Value = rate
                    wsInv.Cells(r, 5).Value = lineTotal
                    subtotal = subtotal + lineTotal
                    r = r + 1
                End If
            Next itemRow

            ' --- totals ---
            taxAmt = subtotal * taxPct / 100
            grandTotal = subtotal + taxAmt
            wsInv.Range("E15").Value = subtotal
            wsInv.Range("E16").Value = taxAmt
            wsInv.Range("E17").Value = grandTotal
        End If
    Next clientRow
End Sub

' Paints the invoice layout (title, labels, item table, totals) on ws.
Private Sub PaintInvoice(ws As Object)
    Dim r As Long, c As Integer

    ws.Range("A1:E1").Merge
    With ws.Range("A1")
        .Value = "TAX INVOICE"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("A3").Value = "Invoice No:"
    ws.Range("A4").Value = "Date:"
    ws.Range("A5").Value = "Bill To:"
    ws.Range("A3:A5").Font.Bold = True
    ws.Range("A3:A5").Font.Color = RGB(31, 78, 121)
    ws.Range("C3:C6").Font.Bold = True

    Dim headers
    headers = Array("#", "Description", "Qty", "Rate (PKR)", "Amount (PKR)")
    For c = 0 To 4
        With ws.Cells(8, c + 1)
            .Value = headers(c)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter
        End With
    Next c

    For r = 9 To 14
        For c = 1 To 5
            ws.Cells(r, c).Borders.LineStyle = xlContinuous
        Next c
        ws.Cells(r, 2).HorizontalAlignment = xlLeft
        ws.Cells(r, 1).HorizontalAlignment = xlCenter
        ws.Cells(r, 3).HorizontalAlignment = xlCenter
        ws.Cells(r, 4).HorizontalAlignment = xlCenter
        ws.Cells(r, 5).HorizontalAlignment = xlCenter
    Next r

    ws.Range("D15").Value = "Subtotal"
    ws.Range("D16").Value = "Tax (16%)"
    ws.Range("D17").Value = "GRAND TOTAL"
    ws.Range("D15:D17").Font.Bold = True
    ws.Range("D15:D17").HorizontalAlignment = xlRight
    ws.Range("E15:E17").Font.Bold = True
    ws.Range("E15:E17").Font.Color = RGB(31, 78, 121)
    ws.Range("E15:E17").NumberFormat = "#,##0"
    ws.Range("E15:E17").HorizontalAlignment = xlCenter

    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 38
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 20
    ws.Range("A19").Value = "Thank you for your business."
    ws.Range("A19").Font.Italic = True
End Sub
