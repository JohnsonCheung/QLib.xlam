Attribute VB_Name = "MDao_BQLin_Write"
Option Explicit
Sub Z_WrtFbqzDb()
Dim P$: P = TmpPth
WrtFbqzDb P, SampDb_DutyDta
BrwPth P
Stop
End Sub

Sub Z_WrtFbqzT()
Dim T$: T = TmpFt
WrtFbqzT T, SampDb_DutyDta, "PermitD"
BrwFt T
End Sub
Sub WrtFbqzDb(Pth, Db As Database)
WrtFbqzTT Pth, Db, Tny(Db)
End Sub

Sub WrtFbqzTT(Pth, Db As Database, TT)
Dim T, P$
P = PthEnsSfx(Pth)
For Each T In TnyzTT(TT)
    WrtFbqzT P & T & ".txt", Db, T
Next
End Sub

Sub WrtFbqzT(Fbq, Db As Database, T)
Dim F%: F = FnoOup(Fbq)
Dim R As Dao.Recordset
Set R = RszT(Db, T)
Dim L$: L = ShtTysColonFldNmBQLinzFds(R.Fields)
Print #F, L
With R
    While Not .EOF
        Print #F, BqlzRs(R)
        .MoveNext
    Wend
    .Close
End With
Close #F
End Sub
Private Function DoczFbq() As String()
Erase XX
X "Fbq is Full file name of back quote (`) separated lines"
X "Fbq has first line as ShtTysColonFldNmQBLin"
X "Fbq rest of lines are records"
DoczFbq = XX
Erase XX
End Function
