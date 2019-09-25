Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "Module1."
Type DiOpt: Som As Boolean: Di As Dictionary: End Type
Function SomDi(D As Dictionary) As DiOpt
SomDi.Som = True
Set SomDi.Di = D
End Function
