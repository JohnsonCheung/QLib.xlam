Attribute VB_Name = "MDao_Att_Op_Exp"
Option Explicit
Const CMod$ = "MDao_Att_Op_Exp."
Public Const DocOfTblAtt$ = ""
Public Const DocOfAtt$ = "Attachment:It a Key-string of Table-Att in a database.  It can retrieve a record from Table-Att."
Private Function ExpAttzAttd$(A As Attd, ToFfn) 'Export the only File in {Attds} {ToFfn}
Const CSub$ = CMod & "ExpAttzAttd"
Dim Fn$, T$, F2 As Dao.Field2
With A.Ars
    If Ext(!Filename) <> Ext(ToFfn) Then Thw CSub, "The Ext in the Att should be same", "Att-Ext ToFfn-Ext", Ext(!Filename), Ext(ToFfn)
    Set F2 = !FileData
End With
F2.SaveToFile ToFfn
ExpAttzAttd = ToFfn
End Function

Function ExpAtt$(Db As Database, Att, ToFfn) 'Exporting the first File in [Att] to [ToFfn]. _
|If no or more than one file in att, error _
|If any, export and return ToFfn
Const CSub$ = CMod & "ExpAtt"
Dim N%
N = AttFilCnt(Db, Att)
If N <> 1 Then
    Thw CSub, "AttNm should have only one file, no export.", _
        "AttNm FilCnt ExpToFile Db", _
        Att, N, ToFfn, DbNm(Db)
End If
ExpAtt = ExpAttzAttd(Attd(Db, Att), ToFfn)
Inf CSub, "Att is exported", "Att ToFfn FmDb", Att, ToFfn, DbNm(Db)
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
Dim Fd2 As Dao.Field2
    Set Fd2 = AttFd2(A, Att, AttFn$)

If IsNothing(Fd2) Then
    Thw CSub, "In record of AttNm there is no given AttFn, but only Act-AttFnAy", _
        "Db Given-AttNm Given-AttFn Act-AttFny ToFfn", _
        DbNm(A), Att, AttFn, AttFnAy(A, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
ExpAttzFn = ToFfn
End Function
Private Function AttFd2(A As Database, Att, AttFn) As Dao.Field2
With Attd(A, Att)
    With .Ars
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

Private Sub ZZ_ExpAtt()
Dim T$, D As Database
T = TmpFx
ExpAttzFn D, "Tp", "TaxRateAlert(Template).xlsm", T
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
ExpAttzFn D, A, A, B
ExpAttzAttd C, B
End Sub

