Attribute VB_Name = "QIde_Ws_Cmp"
Option Explicit
Private Const CMod$ = "MIde_Ws_Cmp."
Private Const Asm$ = "QIde"
Function MdzWs(A As Worksheet) As CodeModule
Set MdzWs = CmpzWs(A).CodeModule
End Function

Function CmpzWs(A As Worksheet) As VBComponent
Set CmpzWs = FstItmzNm(PjzWs(A).VBComponents, A.CodeName)
End Function

Function PjzWs(A As Worksheet) As VBProject
Set PjzWs = WbzWs(A).VBProject
End Function
