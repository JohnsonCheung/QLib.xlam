Attribute VB_Name = "MIde_Ws_Cmp"
Option Explicit
Function MdzWs(A As Worksheet) As CodeModule
Set MdzWs = CmpzWs(A).CodeModule
End Function

Function CmpzWs(A As Worksheet) As VBComponent
Dim P As VBProject
Set P = WbzWs(A).VBProject
Set CmpzWs = FstItrNm(P.VBComponents, A.CodeName)
End Function
