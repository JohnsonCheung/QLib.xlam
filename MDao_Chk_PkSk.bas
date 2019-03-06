Attribute VB_Name = "MDao_Chk_PkSk"
Option Explicit

Function ChkPk$(A As Database, T)
If HasStdPk(A, T) Then Exit Function
If HasPk(A, T) Then
    Dim Pk$(): Pk = PkFny(A, T)
    Select Case True
    Case Sz(Pk) <> 1: ChkPk = FmtQQ("There is PrimaryKey-Idx, but it has [?] fields[?]", Sz(Pk), TLin(Pk))
    Case Pk(0) <> T & "Id": ChkPk = FmtQQ("There is One-field-PrimaryKey-Idx of FldNm(?), but it should named as ?Id", Pk(0), T)
    Case FdzTF(A, T, 0).Name <> T & "Id": ChkPk = FmtQQ("The Pk-field(?Id) should be first fields, but now it is (?)", T, FdzTF(A, T, T & "Id").OrdinalPosition)
    End Select
End If
ChkPk = "[?] does not have PrimaryKey-Idx"
End Function

Function ChkSsk$(A As Database, T)
Dim O$, Sk$(): Sk = SkFny(A, T)
O = ChkSk(A, T): If O <> "" Then ChkSsk = O: Exit Function
If Sz(Sk) <> 1 Then
    ChkSsk = FmtQQ("Secondary is not single field. Tbl[?] Db[?] SkFfn[?]", T, DbNm(A), JnTermAy(Sk))
End If
End Function
Function ChkPkSk(A As Database) As String()
Dim T
For Each T In Tny(A)
    PushIAy ChkPkSk, ChkPkSkzT(A, T)
Next
End Function
Function ChkPkSkzT(A As Database, T) As String()
PushNonBlankStr ChkPkSkzT, ChkPk(A, T)
PushNonBlankStr ChkPkSkzT, ChkSk(A, T)
End Function

Function ChkSk$(A As Database, T)
Dim SkIdx As Dao.Index, I As Dao.Index
If Not HasIdx(A, T, C_SkNm) Then
    ChkSk = FmtQQ("Not SecondaryKey for Table[?] in Db[?]", T, DbNm(A))
    Exit Function
End If
Set SkIdx = A.TableDefs(T).Indexes(C_SkNm)
Select Case True
Case Not SkIdx.Unique
    ChkSk = FmtQQ("SecondaryKey is not unique for Table[?] in Db[?]", T, DbNm(A))
Case Else
    Set I = FstUniqIdx(A, T)
    If Not IsNothing(I) Then
        ChkSk = FmtQQ("No SecondaryKey, but there is uniq idx, it should name as SecondaryKey for Table[?] Db[?] UniqIdxNm[?] IdxFny[?]", _
            T, DbNm(A), I.Name, JnTermAy(FnyzIdx(I)))
    End If
End Select
End Function

