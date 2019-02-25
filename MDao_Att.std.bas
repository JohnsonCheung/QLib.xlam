Attribute VB_Name = "MDao_Att"
Option Explicit
Const CMod$ = "MDaoAtt."
Type Attd
    TRs As DAO.Recordset
    ARs As DAO.Recordset
End Type

Function Attd(Att) As Attd
Attd = Attdz(CDb, Att)
End Function

Function FstAttFfn$(Att)
FstAttFfn = RsMovFst(Attd(Att).ARs)!Filename
End Function

Function AttFilCnt%(Att)
AttFilCnt = AttFilCntz(CDb, Att)
End Function

Function FnyzAttFldDb(A As Database) As String()
Dim TRs As DAO.Recordset2: Set TRs = A.TableDefs("Att").OpenRecordset
Dim ARs As DAO.Recordset2: Set ARs = TRs!Att.Value
FnyzAttFldDb = Itn(ARs.Fields)
End Function

Function HasOneFilAtt(Att) As Boolean
HasOneFilAtt = HasOneFilAttz(CDb, Att)
End Function

Function IsOldAttz(A As Database, Att$, Ffn$) As Boolean
Const CSub$ = CMod & "IsOldAtt"
Dim TAtt As Date, TFfn As Date, AttIs$
TAtt = AttTimz(A, Att)
TFfn = TimFfn(Ffn)
AttIs = IIf(TAtt > TFfn, "new", "old")
Dim M$
M = "Att is " & AttIs
Info CSub, M, "Att Ffn AttTim TimFfn AttIs-Old-or-New?", Att, Ffn, TAtt, TFfn, AttIs
End Function

Function AttSzz&(A As Database, Att)
AttSzz = ValzSskDb(A, "Att", "FilSz", Att)
End Function

Function AttTimz(A As Database, Att) As Date
AttTimz = ValzSskDb(A, "Att", "FilTim", Att)
End Function

Property Get AttNy() As String()
AttNy = AttNyz(CDb)
End Property
Function AttFilCntvRs%(A As Attd)
AttFilCntvRs = NReczRs(A.ARs)
End Function
Function AttFilCntz%(Db As Database, Att)
AttFilCntz = AttFilCntvRs(Attdz(Db, Att))
End Function

Function FnyzAttFld() As String()
FnyzAttFld = FnyzAttFldDb(CDb)
End Function

Function AttFnAy(Att) As String()
AttFnAy = AttFnAyz(CDb, Att)
End Function

Function AttFnAyz(A As Database, Att) As String()
Dim R As Attd: R = Attdz(A, Att)
AttFnAyz = SyzRs(R.ARs, "FileName")
End Function
Function FnyzTblAttDb(A As Database) As String()
FnyzTblAttDb = Fnyz(A, "Att")
End Function
Function FnyzTblAtt() As String()
FnyzTblAtt = FnyzTblAttDb(CDb)
End Function
Function AttFnz(A As Database, Att)
AttFnz = AttFnvRs(Attdz(A, Att))
End Function

Function HasOneFilAttz(A As Database, Att) As Boolean
Debug.Print "DbAttHasOnlyFile: " & Attdz(A, Att).ARs.RecordCount
HasOneFilAttz = Attdz(A, Att).ARs.RecordCount = 1
End Function

Function AttNyz(A As Database) As String()
AttNyz = SyzRs(Rsz(A, "Select AttNm from Att order by AttNm"))
End Function

Private Sub Z_AttFnAy()
'Fb_CDb SampFbzShpRate
D AttFnAy("AA")
'CDb_Cls
End Sub

Private Sub Z()
Z_AttFnAy
MDao_Att_Inf:
End Sub

Function AttNm$(A As Attd)
AttNm = A.TRs!AttNm
End Function

Function AttFn$(Att)
AttFn = AttFnz(CDb, Att)
End Function

Function AttFnvRs$(A As Attd)
Const CSub$ = CMod & "AttFnvRs"
With A.ARs
    If .EOF Then
        If .BOF Then
            Info CSub, "[AttNm] has no attachment files", "AttNm", AttNm(A)
            Exit Function
        End If
    End If
    .MoveFirst
    AttFnvRs = !Filename
End With
End Function

Function Attdz(A As Database, Att) As Attd
With Attdz
    Set .TRs = A.OpenRecordset(FmtQQ("Select Att,FilTim,FilSz from Att where AttNm='?'", Att))
    If .TRs.EOF Then
        A.Execute FmtQQ("Insert into Att (AttNm) values('?')", Att)
        Set .TRs = A.OpenRecordset(FmtQQ("Select Att from Att where AttNm='?'", Att))
    End If
    Set .ARs = .TRs.Fields(0).Value
End With
End Function

