Attribute VB_Name = "MDao_Att"
Option Explicit
Const CMod$ = "MDaoAtt."
Type Attd
    TRs As Dao.Recordset
    ARs As Dao.Recordset
End Type

Function FstAttFfn$(A As Database, Att)
FstAttFfn = RsMovFst(Attd(A, Att).ARs)!Filename
End Function

Function FnyzAttFld(A As Database) As String()
Dim TRs As Dao.Recordset2: Set TRs = A.TableDefs("Att").OpenRecordset
Dim ARs As Dao.Recordset2: Set ARs = TRs!Att.Value
FnyzAttFld = Itn(ARs.Fields)
End Function

Function IsOldAtt(A As Database, Att$, Ffn$) As Boolean
Const CSub$ = CMod & "IsOldAtt"
Dim TAtt As Date, TFfn As Date, AttIs$
TAtt = AttTim(A, Att)
TFfn = TimFfn(Ffn)
AttIs = IIf(TAtt > TFfn, "new", "old")
Dim M$
M = "Att is " & AttIs
Info CSub, M, "Att Ffn AttTim TimFfn AttIs-Old-or-New?", Att, Ffn, TAtt, TFfn, AttIs
End Function

Function AttSz&(A As Database, Att)
AttSz = ValzSsk(A, "Att", "FilSz", Att)
End Function

Function AttTim(A As Database, Att) As Date
AttTim = ValzSsk(A, "Att", "FilTim", Att)
End Function

Function AttFilCntzAttd%(A As Attd)
AttFilCntzAttd = NReczRs(A.ARs)
End Function
Function AttFilCnt%(Db As Database, Att)
AttFilCnt = AttFilCntzAttd(Attd(Db, Att))
End Function

Function AttFnAy(A As Database, Att) As String()
Dim R As Attd: R = Attd(A, Att)
AttFnAy = SyzRs(R.ARs, "FileName")
End Function
Function FnyzTblAtt(A As Database) As String()
FnyzTblAtt = Fny(A, "Att")
End Function
Function AttFn$(A As Database, Att)
AttFn = AttFnzAttd(Attd(A, Att))
End Function

Function HasOneFilAtt(A As Database, Att) As Boolean
Debug.Print "DbAttHasOnlyFile: " & Attd(A, Att).ARs.RecordCount
HasOneFilAtt = Attd(A, Att).ARs.RecordCount = 1
End Function

Function AttNy(A As Database) As String()
AttNy = SyzRs(Rs(A, "Select AttNm from Att order by AttNm"))
End Function

Private Sub Z_AttFnAy()
D AttFnAy(SampDb_ShpCst, "AA")
End Sub

Private Sub Z()
Z_AttFnAy
MDao_Att_Inf:
End Sub

Function AttNm$(A As Attd)
AttNm = A.TRs!AttNm
End Function

Function AttFnzAttd$(A As Attd)
Const CSub$ = CMod & "AttFnzAttd"
With A.ARs
    If .EOF Then
        If .BOF Then
            Info CSub, "[AttNm] has no attachment files", "AttNm", AttNm(A)
            Exit Function
        End If
    End If
    .MoveFirst
    AttFnzAttd = !Filename
End With
End Function

Function Attd(A As Database, Att) As Attd
With Attd
    Set .TRs = A.OpenRecordset(FmtQQ("Select Att,FilTim,FilSz from Att where AttNm='?'", Att))
    If .TRs.EOF Then
        A.Execute FmtQQ("Insert into Att (AttNm) values('?')", Att)
        Set .TRs = A.OpenRecordset(FmtQQ("Select Att from Att where AttNm='?'", Att))
    End If
    Set .ARs = .TRs.Fields(0).Value
End With
End Function

