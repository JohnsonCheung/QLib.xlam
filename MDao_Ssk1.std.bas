Attribute VB_Name = "MDao_Ssk1"
Option Explicit
Public Const C_SkNm$ = "SecondaryKey"
Public Const C_PkNm$ = "PrimaryKey"
Function SkFnyz(A As Database, T) As String()
'If HasSk Then SkFnyz = FnyzIdx(SkIdx)
End Function

Function Sskv() As Aset
'SSskv is [S]ingleFielded [S]econdKey [K]ey [V]alue [Aset], which is always a Value-Aset.
'and Ssk is a field-name from , which assume there is a Unique-Index with name "SecordaryKey" which is unique and and have only one field
'Set Sskv = ColSet(SskFld)
End Function
Function SkIdx(A As Database, T) As DAO.Index
Set SkIdx = Idxz(A, T, C_SkNm)
End Function

Function SkFny(T) As String()

End Function
Function SskFld$(T)

End Function

Function SskFldz$(Db As Database, T)
'Dim Sk$(): Sk = SkFnyz(Db, T): If Sz(Sk) = 1 Then SsFldz = Sk(0)
'Thw CSub, "SkFny-Sz<>1", "Db T, SkFny-Sz SkFny", DbNm(Db), T, Sz(Sk), Sk
End Function

