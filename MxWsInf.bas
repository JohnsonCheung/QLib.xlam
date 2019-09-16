Attribute VB_Name = "MxWsInf"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxWsInf."

Function MdzWs(S As Worksheet) As CodeModule
Set MdzWs = CmpzWs(S).CodeModule
End Function

Function CmpzWs(S As Worksheet) As VBComponent
Set CmpzWs = FstzItn(PjzWs(S).VBComponents, S.CodeName)
End Function

Function PjzWs(S As Worksheet) As VBProject
Set PjzWs = WbzWs(S).VBProject
End Function