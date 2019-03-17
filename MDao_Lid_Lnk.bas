Attribute VB_Name = "MDao_Lid_Lnk"
Option Explicit
Function ErzLnkTblzLtPm(Db As Database, A() As LtPm) As String()
Dim J%
For J = 0 To UB(A)
    With A(J)
        PushIAy ErzLnkTblzLtPm, ErzLnkTblzTSrcCn(Db, .T, .S, .Cn)
    End With
Next
End Function
Sub LnkTblzLtPm(Db As Database, A() As LtPm)
Dim J%
For J = 0 To UB(A)
    With A(J)
        LnkTblzTSCn Db, .T, .S, .Cn
    End With
Next
End Sub
Function TdzTSCn(T, Src, Cn) As Dao.TableDef
Set TdzTSCn = New Dao.TableDef
With TdzTSCn
    .Connect = Cn
    .Name = T
    .SourceTableName = Src
End With
End Function

Sub LnkTblzTSCn(Db As Database, T, S, Cn)
DrpT Db, T
Db.TableDefs.Append TdzTSCn(T, S, Cn)
End Sub

Sub LnkFxw(A As Database, T, Fx, Wsn)
LnkTblzTSCn A, T, Wsn & "$", CnStrzFxDAO(Fx)
End Sub

Sub LnkFbzTT(Db As Database, TTCrt$, Fb$, Optional Fbtt$)
Dim TnyCrt$(), TnyzFb$(), J%, T
TnyCrt = NyzNN(TTCrt)
TnyzFb = IIf(Fbtt = "", TnyCrt, TermAy(Fbtt))
If Si(TnyzFb) <> Si(TnyCrt) Then
    Thw CSub, "[TTCrt] and [FbttSz] are diff", "TTCrtSz FbttSz TnyCrt TnyzFb GivenFbtt", Si(TnyCrt), Si(TnyzFb), TnyCrt, TnyzFb, Fbtt
End If
Dim Cn$: Cn = CnStrzFbDao(Fb)
For J = 0 To UB(TnyCrt)
    LnkTblzTSCn Db, TnyCrt(J), TnyzFb(J), Cn
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
ThwEr ErzLnkTblzTSrcCn(A, T, IIf(Fbt = "", T, Fbt), Cn), CSub
End Sub


