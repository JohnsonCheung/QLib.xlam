Attribute VB_Name = "MDao_Att"
Option Explicit
Const CMod$ = "MDao_Att."
Type Attd
    TRs As DAO.Recordset
    Ars As DAO.Recordset
End Type

Function FstAttFfn$(A As Database, Att)
FstAttFfn = RsMovFst(Attd(A, Att).Ars)!Filename
End Function

Function FnyzAttFld(A As Database) As String()
Dim TRs As DAO.Recordset2: Set TRs = A.TableDefs("Att").OpenRecordset
Dim Ars As DAO.Recordset2: Set Ars = TRs!Att.Value
FnyzAttFld = Itn(Ars.Fields)
End Function

Function IsOldAtt(A As Database, Att$, Ffn$) As Boolean
Const CSub$ = CMod & "IsOldAtt"
Dim TAtt As Date, TFfn As Date, AttIs$
TAtt = AttTim(A, Att)
TFfn = TimFfn(Ffn$)
AttIs = IIf(TAtt > TFfn, "new", "old")
Dim M$
M = "Att is " & AttIs
Inf CSub, M, "Att Ffn AttTim TimFfn AttIs-Old-or-New?", Att, Ffn, TAtt, TFfn, AttIs
End Function

Function AttSz&(A As Database, Att)
AttSz = ValOfSsk(A, "Att", "FilSz", Att)
End Function

Function AttTim(A As Database, Att) As Date
AttTim = ValOfSsk(A, "Att", "FilTim", Att)
End Function

Function AttFilCntzAttd%(A As Attd)
AttFilCntzAttd = NReczRs(A.Ars)
End Function
Function AttFilCnt%(A As Database, Att)
AttFilCnt = AttFilCntzAttd(Attd(Db, Att))
End Function

Function AttFnAy(A As Database, Att) As String()
Dim R As Attd: R = Attd(A, Att)
AttFnAy = SyzRs(R.Ars, "FileName")
End Function
Function FnyzTblAtt(A As Database) As String()
FnyzTblAtt = Fny(A, "Att")
End Function
Function AttFn$(A As Database, Att)
AttFn = AttFnzAttd(Attd(A, Att))
End Function

Function HasOneFilAtt(A As Database, Att) As Boolean
Debug.Print "DbAttHasOnlyFile: " & Attd(A, Att).Ars.RecordCount
HasOneFilAtt = Attd(A, Att).Ars.RecordCount = 1
End Function

Function AttNy(A As Database) As String()
AttNy = SyzRs(Rs(A, "Select AttNm from Att order by AttNm"))
End Function

Private Sub Z_AttFnAy()
D AttFnAy(SampDbzShpCst, "AA")
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
With A.Ars
    If .EOF Then
        If .BOF Then
            Inf CSub, "[AttNm] has no attachment files", "AttNm", AttNm(A)
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
    Set .Ars = .TRs.Fields(0).Value
End With
End Function

