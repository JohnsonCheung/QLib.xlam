Attribute VB_Name = "MDao_BQLin_Read"
Option Explicit
Private Sub Z_CrtTTzPth()
Dim A As Database: Set A = TmpDb
Dim P$: P = TmpPth
WrtFbqzDb P, SampDb_DutyDta
CrtTTzPth A, P
BrwDb A
Stop
End Sub
Sub CrtTTzPth(A As Database, FbqPth)
CrtTTzPthTT A, FbqPth, FnnAy(FbqPth, "*.txt")
End Sub
Sub CrtTTzPthTT(A As Database, FbqPth, TT)
Dim T, P$, Fbq$
P = PthEnsSfx(FbqPth)
For Each T In TnyzTT(TT)
    Fbq = P & T & ".txt"
    CrtTblzFbq A, T, Fbq
Next
End Sub
Private Sub Z_CrtTblzFbq()
Dim Fbq$: Fbq = TmpFt
WrtFbqzT Fbq, SampDb_DutyDta, "PermitD"
Dim D As Database: Set D = TmpDb
CrtTblzFbq D, "PermitD", Fbq
BrwDb D
Stop
End Sub

Sub CrtTblzFbq(A As Database, T, Fbq)
Dim L$, F%, R As Dao.Recordset, J%
F = FnoInp(Fbq)
Line Input #F, L
CrtTblzShtTysColonFldNmBQLin A, T, L
Set R = RszT(A, T)
While Not EOF(F)
    Line Input #F, L
    InsRszBql R, L
    J = J + 1
Wend
Close #F
End Sub

Private Sub CrtTblzShtTysColonFldNmBQLin(A As Database, T, BQLin$)
Dim Td As New Dao.TableDef
Td.Name = T
Dim I
For Each I In Split(BQLin, "`")
    Td.Fields.Append FdzShtTysColonFldNm(I)
Next
A.TableDefs.Append Td
End Sub
Private Function FdzShtTysColonFldNm(A) As Dao.Field
Dim T As Dao.DataTypeEnum
Dim S As Byte
With Brk(A, ":")
    If FstChr(A) = "T" Then
        T = dbText
        S = RmvFstChr(.S1)
    Else
        T = DaoTyzShtTy(.S1)
    End If

    Set FdzShtTysColonFldNm = Fd(.S2, T, TxtSz:=S)
End With
End Function

Function ShtTysColonFldNmBQLinzFds$(A As Dao.Fields)
Dim Ay$(), F As Dao.Field
For Each F In A
    PushI Ay, ShtTyszFd(F) & ":" & F.Name
Next
ShtTysColonFldNmBQLinzFds = Jn(Ay, "`")
End Function

Private Function ShtTyszFd$(A As Dao.Field)
Dim B$: B = ShtTyzDao(A.Type)
If A.Type = dbText Then
    B = B & A.Size
End If
ShtTyszFd = B
End Function
Function DoczShtTysColonFldNm() As String()
Erase XX
X "ShtTysColonFldNmBQLin is BQLin (Back-Quote-Lin) with each term is ShtTyColonFLdNm"
X "ShtTysColonFldNm     is ShtTys + Colon + FldNm"
DoczShtTysColonFldNm = XX
Erase XX
End Function
Function DoczFbq() As String()
Erase XX
X "Fbq is Full file name of back quote (`) separated lines"
X "Fbq fst line is ShtTyColonFldNmBQLin"
X "Fbq is generated from Rs"
DoczFbq = XX
Erase XX
End Function

