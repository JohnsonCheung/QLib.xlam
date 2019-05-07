Attribute VB_Name = "QVb_RRCC"
Option Explicit
Private Const CMod$ = "MVb_RRCC."
Private Const Asm$ = "QVb"
Function RRCC(R1, R2, C1, C2) As RRCC
Dim O As New RRCC
Set RRCC = O.Init(R1, R2, C1, C2)
End Function
