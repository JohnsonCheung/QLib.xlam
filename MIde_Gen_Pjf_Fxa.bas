Attribute VB_Name = "MIde_Gen_Pjf_Fxa"
Option Explicit
Sub Z_DistFxazCompress()
Debug.Print DistFxazCompress(Pjf(CurPj))
End Sub

Function DistFxazCompress$(Fxa)
'ExpFxa Fxa
DistFxazGen SrcpzPjf(Fxa)
DistFxazCompress = Fxa
End Function

Function DistFxazGen$(Srcp, Optional Xls As Excel.Application)
If Not IsSrcp(Srcp) Then Thw CSub, "Not Srcp", "Srcp", Srcp
Dim Fxa$
Dim A As Excel.Application
    Set A = DftXls(Xls)
    Fxa = DistFxa(Srcp)
    Dim Wb As Workbook
    Set Wb = WbCrtNxtFxa(Fxa, A)
DistFxazGen = Wb.FullName
AddRfzPj Wb.VBProject
LoadBas Wb.VBProject
Wb.Close True
If IsNothing(Xls) Then XlsQuit A
End Function

Function PjzFxa(Fxa, A As Excel.Application) As VBProject
If Not IsFxa(Fxa) Then Stop
If Dir(Fxa) = "" Then Stop
Dim O As VBProject: Set O = PjzXls(A, Fxa): If Not IsNothing(O) Then Set PjzFxa = O: Exit Function
A.Workbooks.Open Fxa
Set PjzFxa = PjzXls(A, Fxa)
End Function

Function WbCrtNxtFxa(Fxa, A As Excel.Application) As Workbook
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
Dim Wb As Workbook: Set Wb = A.Workbooks.Add
Dim F$: F = NxtFfn(Fxa)
Wb.SaveAs F, XlFileFormat.xlOpenXMLAddIn
Wb.Close
Set Wb = A.Workbooks.Open(F)
Dim Pj As VBProject
PjzFxa(F, A).Name = Fnn(Fxa)
Wb.Save
Set WbCrtNxtFxa = Wb
End Function

