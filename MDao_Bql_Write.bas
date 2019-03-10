Attribute VB_Name = "MDao_Bql_Write"
Option Explicit
Private Sub Z_WrtFbqlzDb()
Dim P$: P = TmpPth
WrtFbqlzDb P, SampDb_DutyDta
BrwPth P
Stop
End Sub

Private Sub Z_WrtFbqlzT()
Dim T$: T = TmpFt
WrtFbql T, SampDb_DutyDta, "PermitD"
BrwFt T
End Sub

Sub WrtFbqlzDb(Pth, Db As Database)
WrtFbqlzTT Pth, Db, Tny(Db)
End Sub

Sub WrtFbqlzTT(Pth, Db As Database, TT)
Dim T, P$
P = PthEnsSfx(Pth)
For Each T In TnyzTT(TT)
    WrtFbql P & T & ".txt", Db, T
Next
End Sub

Sub WrtFbql(Fbql, Db As Database, T)
Dim F%: F = FnoOup(Fbql)
Dim R As Dao.Recordset
Set R = RszT(Db, T)
Dim L$: L = ShtTyBqlzT(Db, T)
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
X "Fbq has first line as ShtTyscfQBLin"
X "Fbq rest of lines are records"
DoczFbq = XX
Erase XX
End Function
