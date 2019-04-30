Attribute VB_Name = "MIde_Gen_Pjf_Fxa"
Option Explicit
Private Sub Z_CompressFxa()
CompressFxa Pjf(CurPj)
End Sub

Sub CompressFxa(Fxa$, Optional Xls0 As Excel.Application)
PjExp PjzPjfVbe(Xls0.Vbe, Fxa)
Dim Srcp$: Srcp = SrcpzPjf(Fxa)
CrtDistFxa Srcp, Xls0
RplFfn Fxa, Srcp
End Sub

Sub CrtDistFxa(Srcp$, Optional Xls0 As Excel.Application)
If Not IsSrcp(Srcp) Then Thw CSub, "Not Srcp", "Srcp", Srcp
Dim Fxa$
Dim Xls As Excel.Application
    Set Xls = DftXls(Xls0)
    Fxa = DistFxa(Srcp)
    Dim Wb As Workbook
    Set Wb = CrtFxa(Fxa, Xls)
AddRfzPj Wb.VBProject
LoadBas Wb.VBProject
Wb.Close True
If IsNothing(Xls0) Then QuitXls Xls
End Sub

Function PjzFxa(Fxa$, A As Excel.Application) As VBProject
If Not IsFxa(Fxa) Then Stop
If Dir(Fxa) = "" Then Stop
Dim O As VBProject: Set O = PjzXls(A, Fxa): If Not IsNothing(O) Then Set PjzFxa = O: Exit Function
A.Workbooks.Open Fxa
Set PjzFxa = PjzXls(A, Fxa)
End Function

Function CrtFxa(Fxa$, Xls As Excel.Application) As Workbook
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
Dim Wb As Workbook: Set Wb = Xls.Workbooks.Add
PjzFxa(Fxa, Xls).Name = Fnn(Fxa)
Wb.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn
End Function
