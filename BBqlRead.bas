Attribute VB_Name = "BBqlRead"
Option Explicit
Const Asm$ = "Dao"
Const Ns$ = "Dao.Bql"
Const ShtTyBql$ = "Short-Type-Si-Colon-FldNm-Bql:Sht.Ty.s.c.f.Bql: It is a [Bql] with each field is a [ShtTyscf]"
Public Const DocOfBql$ = "Full:Back-Quote-Line.  BasLin:Lin.  Brk:B.q.l.  Back-Quote is (`) and it is a String.  Each field is separated by (`)"
Public Const DocOfFbql$ = "Fullfilename-Bql:F.bql:it is a [Ft]|Each line is a [Bql]|Fst line is [ShtTyBql]"
Public Const DocOfShtTys$ = "ShtTy-Si:It is a [ShtTy] or (Tnnn) where nnn can 1 to 3 digits of value 1-255"
Public Const DocOfShtTyLis$ = "ShtTyLis Short-Type-List Sht.Ty.Lis (String)|is a Cml-String of each 1 to 3 char of ShtTy"
Public Const DocOfShtTyscf$ = "Full: ShtTy-Si-Colon-FldNm. Mmic:ShtTy.s.c.f.  FmTy:ColVal.   If FldNm have space, then ShtTyscf should be sq bracket"
Public Const DocOfShtTyBql$ = "ShtTyscf-Bql:ShtTy.Bql:It is a [Bql] with each field is a [ShtTyscf].  It is used to create an empty table by CrtTblzShtTyscfBql"

Function ShtTyscfBqlzDrs$(A As Drs)
Dim Dry(): Dry = A.Dry
If Si(Dry) = 0 Then ShtTyscfBqlzDrs = Jn(A.Fny, "`"): Exit Function
Dim O$(), F$, I, C&, Fny$()
Fny = A.Fny
For C = 0 To NColzDrs(A) - 1
    F = Fny(C)
    PushI O, ShtTyscfzCol(ColzDry(Dry, C), F)
Next
ShtTyscfBqlzDrs = Jn(O, "`")
End Function

Function ShtTyscfzCol$(Col(), F$)
Dim O$: O = ApdIf(ShtTyszCol(Col), ":") & F
If IsNeedQuote(F) Then O = QuoteSqBkt(O)
ShtTyscfzCol = O
End Function

Private Sub Z_CrtTTzFbqlPth()
Dim A As Database: Set A = TmpDb
Dim P$: P = TmpPth
WrtFbqlzDb P, SampDbzDutyDta
CrtTTzFbqlPth A, P
BrwDb A
End Sub

Sub CrtTTzFbqlPth(A As Database, FbqlPth$)
CrtTTzFbqlPthFnnSy A, FbqlPth, FnnSy(FbqlPth, "*.bql.txt")
End Sub

Sub CrtTTzFbqlPthFnnSy(A As Database, FbqlPth$, FnnSy$())
Dim T, P$, Fbql$
P = EnsPthSfx(FbqlPth)
For Each T In FnnSy
    Fbql = P & T & ".txt"
    CrtTblzFbql A, Fbql
Next
End Sub

Private Sub Z_CrtTblzFbql()
Dim Fbql$: Fbql = TmpFt
WrtFbql Fbql, SampDbzDutyDta, "PermitD"
Dim D As Database: Set D = TmpDb
CrtTblzFbql D, "PermitD", Fbql
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
Set D = Db(Fb$)
For Each IFfn In FfnSy(FbqlPth, "*.bql.txt")
    CrtTblzFbql D, CStr(IFfn)
Next
End Sub

Function TblNmzFbql$(Fbql$)
If Not HasSfx(Fbql, ".bql.txt") Then Thw CSub, "Fbql does not have .bql.txt sfx", "Fbql", Fbql
TblNmzFbql = RmvSfx(Fn(Fbql), ".bql.txt")
End Function

Sub CrtTblzFbql(A As Database, Fbql$, Optional T0$)
Dim T$
    T = T0
    If T = "" Then T = TblNmzFbql(Fbql)

Dim F%, L$, R As Dao.Recordset
F = FnoInp(Fbql)
Line Input #F, L
CrtTblzShtTyscfBql A, T, L

Set R = RszT(A, T)
While Not EOF(F)
    Line Input #F, L
    InsRszBql R, L
Wend
R.Close
Close #F
End Sub

Sub CrtTblzShtTyscfBql(A As Database, T$, ShtTyscfBql$)
Dim Td As New Dao.TableDef
Td.Name = T
Dim I
For Each I In Split(ShtTyscfBql, "`")
    Td.Fields.Append FdzShtTyscf(CStr(I))
Next
A.TableDefs.Append Td
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
    Set FdzShtTyscf = Fd(.S2, T, TxtSz:=S)
End With
End Function

Function ShtTyBqlzT$(A As Database, T$)
Dim Ay$(), F As Dao.Field
For Each F In A.TableDefs(T).Fields
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



