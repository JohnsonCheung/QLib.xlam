Attribute VB_Name = "MDao_Ssk1"
Option Explicit
Public Const C_SkNm$ = "SecondaryKey"
Public Const C_PkNm$ = "PrimaryKey"
Function SkFny(A As Database, T) As String()
'If HasSk Then SkFny = FnyzIdx(SkIdx)
End Function

Function Sskv(A As Database, T) As Aset
'SSskv is [S]ingleFielded [S]econdKey [K]ey [V]alue [Aset], which is always a Value-Aset.
'and Ssk is a field-name from , which assume there is a Unique-Index with name "SecordaryKey" which is unique and and have only one field
'Set Sskv = ColSet(SskFld)
End Function
Function SkIdx(A As Database, T) As Dao.Index
Set SkIdx = Idx(A, T, C_SkNm)
End Function

Function SskFld$(Db As Database, T)
'Dim Sk$(): Sk = SkFny(Db, T): If Sz(Sk) = 1 Then SsFldz = Sk(0)
'Thw CSub, "SkFny-Sz<>1", "Db T, SkFny-Sz SkFny", DbNm(Db), T, Sz(Sk), Sk
End Function

