Attribute VB_Name = "MDao_Bql_Read"
Option Explicit
Private Sub Z_CrtTTzPth()
Dim A As Database: Set A = TmpDb
Dim P$: P = TmpPth
WrtFbqlzDb P, SampDb_DutyDta
CrtTTzPth A, P
BrwDb A
Stop
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

Sub CrtTblzFbql(A As Database, T, Fbq)
Dim L$, F%, R As Dao.Recordset, J%
F = FnoInp(Fbq)
Line Input #F, L
CrtTblzShtTyscFldNmBql A, T, L
Set R = RszT(A, T)
While Not EOF(F)
    Line Input #F, L
    InsRszBql R, L
    J = J + 1
Wend
Close #F
End Sub

Sub CrtTblzShtTyscFldNmBql(A As Database, T, ShtTyscFldNmBql$)
Dim Td As New Dao.TableDef
Td.Name = T
Dim I
For Each I In Split(ShtTyscFldNmBql, "`")
    Td.Fields.Append FdzShtTyscFldNm(I)
Next
A.TableDefs.Append Td
End Sub

Private Function FdzShtTyscFldNm(A) As Dao.Field
Dim T As Dao.DataTypeEnum
Dim S As Byte
With Brk2(A, ":")
    Select Case True
    Case .S1 = "":          T = dbText: S = 255
    Case FstChr(A) = "T":   T = dbText: S = RmvFstChr(.S1)
    Case Else:              T = DaoTyzShtTy(.S1)
    End Select
    Set FdzShtTyscFldNm = Fd(.S2, T, TxtSz:=S)
End With
End Function
Function DoczBql() As String()
Erase XX
X "Bql Back Quote (`) Line (String)"
X ". Each field is separated by (`)"
DoczBql = XX
Erase XX
End Function
Function DoczShtTyscFldNmBql() As String()
Erase XX
X "ShtTyscFldNmBql Short-Type-Sz-Colon-FldNm-Bql (String)"
X ". It can used to create an empty table"
DoczShtTyscFldNmBql = XX
Erase XX
End Function

Function ShtTyscFldNmBqlzT$(A As Database, T)
Dim Ay$(), F As Dao.Field
For Each F In A.TableDefs(T).Fields
    PushI Ay, ShtTyszFd(F) & ":" & F.Name
Next
ShtTyscFldNmBqlzT = Jn(Ay, "`")
End Function

Private Function ShtTyszFd$(A As Dao.Field)
Dim B$: B = ShtTyzDao(A.Type)
If A.Type = dbText Then
    B = B & A.Size
End If
ShtTyszFd = B
End Function
Function DoczShtTyLis() As String()
Erase XX
X "ShtTyLis Short-Type-List Sht.Ty.Lis (String)"
X ". is a Cml-String of each 1 to 3 char of ShtTy"
DoczShtTyLis = XX
Erase XX
End Function

Function DoczShtTy() As String()
Erase XX
X "ShtTys Short-Type Sht.Ty (String)"
X ". is a 1 to 3 char as described in ShtTyLis"
DoczShtTy = XX
Erase XX
End Function

Function DoczShtTys() As String()
Erase XX
X "ShtTys Short-Type-Size Sht.Ty.s (String)"
X ". is a ShtTy variant, when it is first char is T following can 1 to 3 digits of TxtSz (max 255)"
DoczShtTys = XX
Erase XX
End Function

Function DoczShtTyscFldNm() As String()
Erase XX
X "ShtTyscFldNm Sht.Short Ty.Type s.sz c.colon FldNm"
X ". FldNm can have space, then ShtTyscFldNm should be sq bracket"
DoczShtTyscFldNm = XX
Erase XX
End Function

Function DoczFbql() As String()
Erase XX
X "Bql  is b.ack q.uote (`) separated l.ines"
X "Fbql is F.ull file name of b.ack q.uote (`) separated l.ines"
X "Fbql fst line is ShtTyscFldNmBqlzT"
X "Fbql is generated from Rs"
DoczFbql = XX
Erase XX
End Function

