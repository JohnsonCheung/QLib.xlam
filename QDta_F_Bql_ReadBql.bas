Attribute VB_Name = "QDta_F_Bql_ReadBql"
Option Compare Text
Option Explicit
Private Const CMod$ = "BBqlRead."
Const Asm$ = "Dao"
Const Ns$ = "Dao.Bql"
Const ShtTyBql$ = "Short-Type-Si-Colon-FldNm-Bql:Sht.Ty.s.c.f.Bql: It is a [Bql] with each field is a [ShtTyscf]"
':Bql: :Lin #Back-Qte-Line# ! Back-Qte is (`) and it is a String.  Each field is separated by (`)
':Fbql: :Ft #Fullfilename-Bql# ! Each line is a [Bql]|Fst line is [ShtTyBql]
':ShtTys: ShtTy-Si:It is a [ShtTy] or (Tnnn) where nnn can 1 to 3 digits of value 1-255"
':ShtTyLis: :Cml #Short-Type-List# ! Each :Cml is 1 or 3 chr of :ShtTy
':ShtTyscf: :Term #ShtTy-Si-Colon-FldNm#  ! If FldNm have space, then ShtTyscf should be sq bracket"
':ShtTyBql: :Bql #ShtTyscf-Bql# ! Each field is a [ShtTyscf].  It is used to create an empty table by CrtTzShtTyscfBql"

Function ShtTyscfBqlzDrs$(A As Drs)
Dim Dy(): Dy = A.Dy
If Si(Dy) = 0 Then ShtTyscfBqlzDrs = Jn(A.Fny, "`"): Exit Function
Dim O$(), F$, I, C&, Fny$()
Fny = A.Fny
For C = 0 To NColzDrs(A) - 1
    F = Fny(C)
    PushI O, ShtTyscfzCol(ColzDy(Dy, C), F)
Next
ShtTyscfBqlzDrs = Jn(O, "`")
End Function

Function ShtTyscfzCol$(Col(), F$)
Dim O$: O = ApdIf(ShtTyszCol(Col), ":") & F
If IsNeedQte(F) Then O = QteSqBkt(O)
ShtTyscfzCol = O
End Function

Private Sub Z_CrtTTzFbqlPth()
Dim D As Database: Set D = TmpDb
Dim P$: P = TmpPth
WrtFbqlzDb P, SampDboDutyDta
CrtTTzFbqlPth D, P
BrwDb D
End Sub

Sub CrtTTzFbqlPth(D As Database, FbqlPth$)
CrtTTzFbqlPthFnny D, FbqlPth, Fnny(FbqlPth, "*.bql.txt")
End Sub

Sub CrtTTzFbqlPthFnny(D As Database, FbqlPth$, Fnny$())
Dim T, P$, Fbql$
P = EnsPthSfx(FbqlPth)
For Each T In Fnny
    Fbql = P & T & ".txt"
    CrtTzFbql D, Fbql
Next
End Sub

Private Sub Z_CrtTzFbql()
Dim Fbql$: Fbql = TmpFt
WrtFbql Fbql, SampDboDutyDta, "PermitD"
Dim D As Database: Set D = TmpDb
CrtTzFbql D, "PermitD", Fbql
BrwDb D
Stop
End Sub

Sub CrtFbzBqlPth(FbqlPth$, Optional Fb0$)
Dim Fb$
    Fb = Fb0
    If Fb = "" Then Fb = FbqlPth & Fdr(FbqlPth) & ".accdb"
DltFfnIf Fb
CrtFb Fb
Dim D As Database, IFfn, T$
Set D = Db(Fb)
For Each IFfn In Ffny(FbqlPth, "*.bql.txt")
    CrtTzFbql D, CStr(IFfn)
Next
End Sub

Function TblNmzFbql$(Fbql$)
If Not HasSfx(Fbql, ".bql.txt") Then Thw CSub, "Fbql does not have .bql.txt sfx", "Fbql", Fbql
TblNmzFbql = RmvSfx(Fn(Fbql), ".bql.txt")
End Function

Sub CrtTzFbql(D As Database, Fbql$, Optional T0$)
Dim T$
    T = T0
    If T = "" Then T = TblNmzFbql(Fbql)

Dim F%, L$, R As Dao.Recordset
F = FnoI(Fbql)
Line Input #F, L
CrtTzShtTyscfBql D, T, L

Set R = RszT(D, T)
While Not EOF(F)
    Line Input #F, L
    InsRszBql R, L
Wend
R.Close
Close #F
End Sub

Sub CrtTzShtTyscfBql(D As Database, T, ShtTyscfBql$)
Dim Td As New Dao.TableDef
Td.Name = T
Dim I
For Each I In Split(ShtTyscfBql, "`")
    Td.Fields.Append FdzShtTyscf(CStr(I))
Next
D.TableDefs.Append Td
End Sub

Private Function FdzShtTyscf(ShtTyscf$) As Dao.Field
Dim T As Dao.DataTypeEnum
Dim S As Byte
With Brk2(ShtTyscf, ":")
    Select Case True
    Case .S1 = "":                 T = dbText: S = 255
    Case FstChr(ShtTyscf) = "T":   T = dbText: S = RmvFstChr(.S1)
    Case Else:                     T = DaoTyzShtTy(.S1)
    End Select
    Dim ZLen As Boolean: ZLen = T = dbText
    Set FdzShtTyscf = Fd(.S2, T, TxtSz:=S, ZLen:=ZLen)
End With
End Function

Function ShtTyBqlzT$(D As Database, T)
Dim Ay$(), F As Dao.Field
For Each F In D.TableDefs(T).Fields
    PushI Ay, ShtTyszFd(F) & ":" & F.Name
Next
ShtTyBqlzT = Jn(Ay, "`")
End Function

Private Function ShtTyszFd$(A As Dao.Field)
Dim B$: B = ShtTyzDao(A.Type)
If A.Type = dbText Then
    B = B & A.Size
End If
ShtTyszFd = B
End Function


