Attribute VB_Name = "MIde_Gen_Pjf_Fxa"
Option Explicit
Sub GenFxa(SrcPth, Optional Xls As Excel.Application)
ThwNoPth SrcPth, CSub
Dim Fxa$: Fxa = DistFxa(SrcPth)
Dim A As Excel.Application: Set A = DftXls(Xls)
Dim Wb As Workbook: Set Wb = CrtFxa(Fxa, A)
Dim Pj As VBProject: Set Pj = PjzFxa(Fxa, A)
AddRfzPj Pj
LoadBas Pj
Wb.Close True
CpyFilzToPth Fxa, AddFdrEns(ParPth(ParPth(Pth(Fxa))), "Dist"), OvrWrt:=True
If IsNothing(Xls) Then XlsQuit A
End Sub

Function PjzFxa(Fxa, A As Excel.Application) As VBProject
If Not IsFxa(Fxa) Then Stop
If Dir(Fxa) = "" Then Stop
Dim O As VBProject: Set O = PjzXls(A, Fxa): If Not IsNothing(O) Then Set PjzFxa = O: Exit Function
A.Workbooks.Open Fxa
Set PjzFxa = PjzXls(A, Fxa)
End Function

Function CrtFxa(Fxa, A As Excel.Application) As Workbook
If Not IsFxa(Fxa) Then Stop
Dim Wb As Workbook: Set Wb = A.Workbooks.Add
Wb.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn
Wb.Close
Set Wb = A.Workbooks.Open(Fxa)
Dim Pj As VBProject
PjzFxa(Fxa, A).Name = Fnn(Fxa)
Wb.Save
Set CrtFxa = Wb
End Function

