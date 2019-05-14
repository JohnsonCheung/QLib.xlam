Attribute VB_Name = "QXls_Wb_WbOp"
Sub DltSheet1(Wb As Workbook)
If Wb.Sheets.Count = 1 Then Exit Sub
If HasWs(Wb, "Sheet") Then WszWb(Wb, "Sheet1").Delete
End Sub
