Attribute VB_Name = "MDao_Bql_Read"
Option Explicit
Public Const DocOfBql$ = "Back-Quote-Line:B.q.l:Back-Quote is (`) and it is a String.  Each field is separated by (`)"
Public Const DocOfFbql$ = "Fullfilename-Bql:F.bql:it is a [Ft]|Each line is a [Bql]|Fst line is [ShtTyBql]"
Public Const DocOfShtTys$ = "ShtTy-Si:It is a [ShtTy] or (Tnnn) where nnn can 1 to 3 digits of value 1-255"
Public Const DocOfShtTyLis$ = "ShtTyLis Short-Type-List Sht.Ty.Lis (String)|is a Cml-String of each 1 to 3 char of ShtTy"
Public Const DocOfShtTyscf$ = "ShtTys-Colon-FldNm:ShtTys.c.f:FldNm can have space, then ShtTyscf should be sq bracket"
Public Const DocOfShtTyBql$ = "ShtTyscf-Bql:ShtTy.Bql:It is a [Bql] with each field is a [ShtTyscf].  It is used to create an empty table by CrtTblzShtTyBql"

Private Sub Z_CrtTTzPth()
Dim A As Database: Set A = TmpDb
Dim P$: P = TmpPth
WrtFbqlzDb P, SampDb_DutyDta
CrtTTzPth A, P
BrwDb A
End Sub

Sub CrtTTzPth(A As Database, FbqlPth)
CrtTTzPthTT A, FbqlPth, FnnAy(FbqlPth, "*.txt")
End Sub

Sub CrtTTzPthTT(A As Database, FbqlPth, TT)
Dim T, P$, Fbql$
P = PthEnsSfx(FbqlPth)
For Each T In TnyzTT(TT)
    Fbql = P & T & ".txt"
    CrtTblzFbql A, T, Fbql
Next
End Sub

Private Sub Z_CrtTblzFbql()
Dim Fbql$: Fbql = TmpFt
WrtFbql Fbql, SampDb_DutyDta, "PermitD"
Dim D As Database: Set D = TmpDb
CrtTblzFbql D, "PermitD", Fbql
BrwDb D
Stop
End Sub

Sub CrtFbzBqlPth(BqlPth, Optional Fb0$)
Dim Fb$
If Fb0 = "" Then
    Fb = BqlPth & Fdr(BqlPth) & ".accdb"
Else
    Fb = Fb0
End If
DltFfnIf Fb
CrtFb Fb
Dim D As Database, Ffn, T$
Set D = Db(Fb)
For Each Ffn In FfnAy(BqlPth, "*.bql.txt")
    CrtTblzFbql D, Fnn(Fnn(Ffn)), Ffn
Next
End Sub

Sub CrtTblzFbql(A As Database, T, Fbq)
Dim L$, F%, R As Dao.Recordset, J%
F = FnoInp(Fbq)
Line Input #F, L
CrtTblzShtTyBql A, T, L
Set R = RszT(A, T)
While Not EOF(F)
    Line Input #F, L
    InsRszBql R, L
    J = J + 1
Wend
Close #F
End Sub

Sub CrtTblzShtTyBql(A As Database, T, ShtTyBql$)
Dim Td As New Dao.TableDef
Td.Name = T
Dim I
For Each I In Split(ShtTyBql, "`")
    Td.Fields.Append FdzShtTyscf(I)
Next
A.TableDefs.Append Td
End Sub

Private Function FdzShtTyscf(A) As Dao.Field
Dim T As Dao.DataTypeEnum
Dim S As Byte
With Brk2(A, ":")
    Select Case True
    Case .S1 = "":          T = dbText: S = 255
    Case FstChr(A) = "T":   T = dbText: S = RmvFstChr(.S1)
    Case Else:              T = DaoTyzShtTy(.S1)
    End Select
    Set FdzShtTyscf = Fd(.S2, T, TxtSz:=S)
End With
End Function

Function ShtTyBqlzT$(A As Database, T)
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



