Attribute VB_Name = "MDao_Chk_PkSk"
Option Explicit
Property Get ChkPkz$(A As Database, T)
Dim S%, K$(), O
K = FnyzIdx(PkIdxz(A, T))
S = Sz(K)
Select Case True
Case S = 0: O = FmtQQ("T[?] has no Pk", T)
Case S = 1:
    If K(0) <> T & "Id" Then
        O = FmtQQ("T[?] has 1-field-Pk of Fld[?].  It should be [?Id].  Db[?]", T, K(0), T, DbNm(A))
    End If
Case Else
    O = FmtQQ("T[?] has primary key.  It should have single field and name eq to table, but now it has Pk[?].  Db[?]", T, JnSpc(K), DbNm(A))
End Select
ChkPkz = O
End Property

Function ChkSskz$(A As Database, T)
Dim O$, Sk$(): Sk = SkFnyz(A, T)
O = ChkSkz(A, T): If O <> "" Then ChkSskz = O: Exit Function
If Sz(Sk) <> 1 Then
    ChkSskz = FmtQQ("Secondary is not single field. Tbl[?] Db[?] SkFfn[?]", T, DbNm(A), JnTermAy(Sk))
End If
End Function
Function ChkPkSkz(A As Database) As String()
Dim T
For Each T In Tnyz(A)
    PushIAy ChkPkSkz, ChkPkSkzT(A, T)
Next
End Function
Function ChkPkSkzT(A As Database, T) As String()
PushNonBlankStr ChkPkSkzT, ChkPkz(A, T)
PushNonBlankStr ChkPkSkzT, ChkSkz(A, T)
End Function

Function ChkSkz$(A As Database, T)
Dim SkIdx As DAO.Index, I As DAO.Index
If Not HasIdxz(A, T, C_SkNm) Then
    ChkSkz = FmtQQ("Not SecondaryKey for Table[?] in Db[?]", T, DbNm(A))
    Exit Function
End If
Set SkIdx = A.TableDefs(T).Indexes(C_SkNm)
Select Case True
Case Not SkIdx.Unique
    ChkSkz = FmtQQ("SecondaryKey is not unique for Table[?] in Db[?]", T, DbNm(A))
Case Else
    Set I = FstUniqIdxz(A, T)
    If Not IsNothing(I) Then
        ChkSkz = FmtQQ("No SecondaryKey, but there is uniq idx, it should name as SecondaryKey for Table[?] Db[?] UniqIdxNm[?] IdxFny[?]", _
            T, DbNm(A), I.Name, JnTermAy(FnyzIdx(I)))
    End If
End Select
End Function

