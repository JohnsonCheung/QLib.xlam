Attribute VB_Name = "MxPutAy"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxPutAy."
Sub PutAyH(AyH, At As Range)
Put_Sq_At SqH(AyH), At
End Sub

Sub PutAyV(AyV, At As Range)
Put_Sq_At SqV(AyV), At
End Sub
Sub PutSSH(SSH$, At As Range)
PutAyH SyzSS(SSH), At
End Sub

Sub PutSSV(SSV$, At As Range)
PutAyV SyzSS(SSV), At
End Sub


