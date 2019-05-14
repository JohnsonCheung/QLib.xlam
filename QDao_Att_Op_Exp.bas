Attribute VB_Name = "QDao_Att_Op_Exp"
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Att_Op_Exp."
Public Const DoczTblAtt$ = ""
Public Const DoczAtt$ = "Attachment:It a Key-string of Table-Att in a database.  It can retrieve a record from Table-Att."
Private Function ExpAttzAttd$(A As Attd, ToFfn$) 'Export the only File in {Attds} {ToFfn}
Const CSub$ = CMod & "ExpAttzAttd"
Dim Fn$, T$, F2 As Dao.Field2
With A.Ars
    If Ext(!Filename) <> Ext(ToFfn) Then Thw CSub, "The Ext in the Att should be same", "Att-Ext ToFfn-Ext", Ext(!Filename), Ext(ToFfn)
    Set F2 = !FileData
End With
F2.SaveToFile ToFfn
ExpAttzAttd = ToFfn
End Function

Function ExpAtt$(A As Database, Att$, ToFfn$) 'Exporting the first File in [Att] to [ToFfn] if Att is newer or ToFfn not exist. _
Er if no or more than one file in att, error. _
Er if any, export and return ToFfn.
Const CSub$ = CMod & "ExpAtt"
Dim N%
N = AttFilCnt(A, Att)
If N <> 1 Then
    Thw CSub, "AttNm should have only one file, no export.", _
        "AttNm FilCnt ExpToFile A", _
        Att, N, ToFfn, Dbn(A)
End If
ExpAtt = ExpAttzAttd(Attd(A, Att), ToFfn)
Inf CSub, "Att is exported", "Att ToFfn FmDb", Att, ToFfn, Dbn(A)
End Function

Function ExpAttzFn$(A As Database, Att$, AttFn$, ToFfn$)
Const CSub$ = CMod & "ExpAttzFn"
If Ext(AttFn) <> Ext(ToFfn) Then
    Thw CSub, "AttFn & ToFfn are dif extEnsion|" & _
        "To export an AttFn to ToFfn, their file extEnsion should be same", _
        "AttFn-Ext ToFfn-Ext A AttNm AttFn ToFfn", _
        Ext(AttFn), Ext(ToFfn), Dbn(A), Att, AttFn, ToFfn
End If
If HasFfn(ToFfn) Then
    Thw CSub, "ToFfn Has, no over write", _
        "A AttNm AttFn ToFfn", _
        Dbn(A), Att, AttFn, ToFfn
End If
Dim Fd2 As Dao.Field2
    Set Fd2 = AttFd2(A, Att, AttFn$)

If IsNothing(Fd2) Then
    Thw CSub, "In record of AttNm there is no given AttFn, but only Act-AttFnAy", _
        "A Given-AttNm Given-AttFn Act-AttFny ToFfn", _
        Dbn(A), Att, AttFn, AttFnAy(A, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
ExpAttzFn = ToFfn
End Function

Private Function AttFd2(A As Database, Att$, AttFn$) As Dao.Field2
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

Private Sub ZZ()
End Sub

