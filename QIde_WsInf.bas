Attribute VB_Name = "QIde_WsInf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ws_Cmp."
Private Const Asm$ = "QIde"

Function MdzWs(S As Worksheet) As CodeModule
Set MdzWs = CmpzWs(S).CodeModule
End Function

Function CmpzWs(S As Worksheet) As VBComponent
Set CmpzWs = FstzItn(PjzWs(S).VBComponents, S.CodeName)
End Function

Function PjzWs(S As Worksheet) As VBProject
Set PjzWs = WbzWs(S).VBProject
End Function
