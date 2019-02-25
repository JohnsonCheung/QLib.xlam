Attribute VB_Name = "MDao_Lid_Lnk"
Option Explicit
Function ChkLnkTbl(Db As Database, A() As LtPm) As String()
Dim J%
For J = 0 To UB(A)
    With A(J)
        PushIAy ChkLnkTbl, ChkLnkTblzDbtSrcCn(Db, .T, .S, .Cn)
    End With
Next
End Function
Sub LnkTblz(Db As Database, A() As LtPm)
Dim J%
For J = 0 To UB(A)
    With A(J)
        LnkTblzDbtSrcCn Db, .T, .S, .Cn
    End With
Next
End Sub
Private Function LtPmAyFx(A() As LiFx, FfnDic As Dictionary) As LtPm()
Dim J%, Fx$, M As LtPm
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        Fx = FfnDic(.Fxn)
        PushObj LtPmAyFx, M.Init(">" & .T, .Wsn & "$", CnStrzFxDAO(Fx))
    End With
Next
End Function
Private Function LtPmAyFb(A() As LiFb, FfnDic As Dictionary) As LtPm()
Dim J%, Fb$, M As LtPm
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        Fb = FfnDic(.Fbn)
        PushObj LtPmAyFb, M.Init(">" & .T, .T, CnStrzFxAdo(Fb))
    End With
Next
End Function

Function LtPm(A As LiPm) As LtPm()
Dim O() As LtPm, D As Dictionary
Set D = A.FilNmToFfnDic
PushObjAy O, LtPmAyFb(A.Fb, D)
PushObjAy O, LtPmAyFx(A.Fx, D)
LtPm = O
End Function
Function TdTSrcCn(T, Src, Cn) As DAO.TableDef
Set TdTSrcCn = New DAO.TableDef
With TdTSrcCn
    .Connect = Cn
    .Name = T
    .SourceTableName = Src
End With
End Function
Sub LnkTblzDbtSrcCn(Db As Database, T, S, Cn)
Drpz Db, T
Db.TableDefs.Append TdTSrcCn(T, S, Cn)
End Sub

Sub LnkWsz(A As Database, T, Fx, Wsn)
LnkTblzDbtSrcCn A, T, Wsn & "$", CnStrzFxDAO(Fx)
End Sub

Sub LnkFbztt(Db As Database, TTCrt$, Fb$, Optional Fbtt$)
Dim TnyCrt$(), TnyzFb$(), J%, T
TnyCrt = FnyzFF(TTCrt)
TnyzFb = IIf(Fbtt = "", TnyCrt, TermAy(Fbtt))
If Sz(TnyzFb) <> Sz(TnyCrt) Then
    Thw CSub, "[TTCrt] and [FbttSz] are diff", "TTCrtSz FbttSz TnyCrt TnyzFb GivenFbtt", Sz(TnyCrt), Sz(TnyzFb), TnyCrt, TnyzFb, Fbtt
End If
Dim Cn$: Cn = CnStrzFbDao(Fb)
For J = 0 To UB(TnyCrt)
    LnkTblzDbtSrcCn Db, TnyCrt(J), TnyzFb(J), Cn
Next
End Sub

Function LnkTnyDb(Db As Database) As String()
Dim T As TableDef
For Each T In Db.TableDefs
    If T.Connect <> "" Then
        PushI LnkTnyDb, T.Name
    End If
Next
End Function

Sub LnkFb(A As Database, T, Fb$, Optional Fbt)
Dim Cn$: Cn = CnStrzFbDao(Fb)
ThwEr ChkLnkTblzDbtSrcCn(A, T, IIf(Fbt = "", T, Fbt), Cn), CSub
End Sub


