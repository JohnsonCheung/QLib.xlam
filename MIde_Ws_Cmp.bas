Attribute VB_Name = "MIde_Ws_Cmp"
Option Explicit
Function MdzWs(A As Worksheet) As CodeModule
Set MdzWs = CmpzWs(A).CodeModule
End Function

Function CmpzWs(A As Worksheet) As VBComponent
Set CmpzWs = FstItmzNm(PjzWs(A).VBComponents, A.CodeName)
End Function

Function PjzWs(A As Worksheet) As VBProject
Set PjzWs = WbzWs(A).VBProject
End Function
