Attribute VB_Name = "QDao_Db_Get_Col"
Option Explicit
Private Const CMod$ = "MDao_Db_Get_Col."
Private Const Asm$ = "QDao"

Function IntAyzQ(A As Database, Q) As Integer()
End Function

Function SyzTF(A As Database, T, F$) As String()
SyzTF = SyzRs(RszTF(A, T, F))
End Function

Function IntozTF(Into, A As Database, T, F$)
IntozTF = IntozRs(Into, RszTF(A, T, F))
End Function
