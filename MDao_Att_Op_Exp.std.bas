Attribute VB_Name = "MDao_Att_Op_Exp"
Option Explicit
Const CMod$ = "MDao_Att_Op_Exp."

Function ExpAtt$(Att, ToFfn)
'Exporting the only file in Att & Return ToFfn
ExpAtt = ExpAttz(CDb, Att, ToFfn)
End Function

Function ExpAttvFn$(Att$, AttFn$, ToFfn)
ExpAttvFn = ExpAttzFn(CDb, Att, AttFn, ToFfn)
End Function

Private Function ExpAttvRs$(A As Attd, ToFfn)
'Export the only File in {Attds} {ToFfn}
Dim Fn$, T$, F2 As DAO.Field2
With A.ARs
    If Ext(!Filename) <> Ext(ToFfn) Then Thw CSub, "The Ext in the Att should be same", "Att-Ext ToFfn-Ext", Ext(!Filename), Ext(ToFfn)
    Set F2 = !FileData
End With
F2.SaveToFile ToFfn
ExpAttvRs = ToFfn
End Function

Function ExpAttz$(Db As Database, Att, ToFfn)
'Exporting the first File in Att.
'If no or more than one file in att, error
'If any, export and return ToFfn
Const CSub$ = CMod & "ExptAttz"
Dim N%
N = AttFilCntz(Db, Att)
If N <> 1 Then
    Thw CSub, "AttNm should have only one file, no export.", _
        "AttNm FilCnt ExpToFile Db", _
        Att, N, ToFfn, DbNm(Db)
End If
ExpAttz = ExpAttvRs(Attdz(Db, Att), ToFfn)
Info CSub, "Att is exported", "Att ToFfn FmDb", Att, ToFfn, DbNm(Db)
End Function

Function ExpAttzFn$(A As Database, Att$, AttFn$, ToFfn)
Const CSub$ = CMod & "ExpAttzFn"
If Ext(AttFn) <> Ext(ToFfn) Then
    Thw CSub, "AttFn & ToFfn are dif extEnsion|" & _
        "To export an AttFn to ToFfn, their file extEnsion should be same", _
        "AttFn-Ext ToFfn-Ext Db AttNm AttFn ToFfn", _
        Ext(AttFn), Ext(ToFfn), DbNm(A), Att, AttFn, ToFfn
End If
If HasFfn(ToFfn) Then
    Thw CSub, "ToFfn Has, no over write", _
        "Db AttNm AttFn ToFfn", _
        DbNm(A), Att, AttFn, ToFfn
End If
Dim Fd2 As DAO.Field2
    Set Fd2 = AttFd2(A, Att, AttFn$)

If IsNothing(Fd2) Then
    Thw CSub, "In record of AttNm there is no given AttFn, but only Act-AttFnAy", _
        "Db Given-AttNm Given-AttFn Act-AttFny ToFfn", _
        DbNm(A), Att, AttFn, AttFnAyz(A, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
ExpAttzFn = ToFfn
End Function
Private Function AttFd2(A As Database, Att, AttFn) As DAO.Field2
With Attdz(A, Att)
    With .ARs
        .MoveFirst
        While Not .EOF
            If !Filename = AttFn Then
                Set AttFd2 = !FileData
            End If
            .MoveNext
        Wend
    End With
End With
End Function

Private Sub ZZ_ExpAttz()
Dim T$
T = TmpFx
ExpAttzFn CDb, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert HasFfn(T)
Kill T
End Sub

Private Sub Z()
End Sub

Private Sub ZZ()
Dim A$
Dim B
Dim C As Attd
Dim D As Database
Dim XX
ExpAttvFn A, A, B
ExpAttvRs C, B
End Sub
