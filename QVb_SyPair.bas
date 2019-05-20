Attribute VB_Name = "QVb_SyPair"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_SyPair."
Type Syab
    A() As String
    B() As String
End Type
Function Syab(A$(), B$()) As Syab
Syab.A = A
Syab.B = B
End Function
