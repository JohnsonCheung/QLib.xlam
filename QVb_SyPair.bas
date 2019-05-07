Attribute VB_Name = "QVb_SyPair"
Option Explicit
Private Const CMod$ = "MVb_SyPair."
Private Const Asm$ = "QVb"
Function SyPair(A, B) As SyPair
Dim O As New SyPair
Set SyPair = O.Init(A, B)
End Function
