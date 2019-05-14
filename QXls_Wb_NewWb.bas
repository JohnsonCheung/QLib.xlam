Attribute VB_Name = "QXls_Wb_NewWb"
Function WbzDs(A As Ds) As Workbook
Dim O As Workbook
Set O = NewWb
With FstWs(O)
   .Name = "Ds"
   .Range("A1").Value = A.DsNm
End With
Dim J%, Ay() As Dt
For J = 0 To A.N - 1
    'WszWbDt O, Ay(J)
Next
Set WbzDs = O
End Function

