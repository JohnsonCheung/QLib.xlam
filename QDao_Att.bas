Attribute VB_Name = "QDao_Att"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Att."
Type Attd
    TblRs As Dao.Recordset '..Att.. #Tbl-Rs ! It is the Tbl-Att recordset
    AttRs As Dao.Recordset '.       #Att-Rs2 !
End Type
Public Const DoczTblAtt$ = ""
Public Const DoczAtt$ = "Attachment:It a Key-string of Table-Att in a database.  It can retrieve a record from Table-Att."

Private Function ExpAttzAttd$(A As Attd, ToFfn$) 'Export the only File in {Attds} {ToFfn}
Const CSub$ = CMod & "ExpAttzAttd"
Dim Fn$, T$, F2 As Dao.Field2
With A.AttRs
    If Ext(!Filename) <> Ext(ToFfn) Then Thw CSub, "The Ext in the Att should be same", "Att-Ext ToFfn-Ext", Ext(!Filename), Ext(ToFfn)
    Set F2 = !FileData
End With
F2.SaveToFile ToFfn
ExpAttzAttd = ToFfn
End Function

Function ExpAtt$(D As Database, Att$, ToFfn$)
'Ret Exporting the first File in [Att] to [ToFfn] if Att is newer or ToFfn not exist.
'Er if no or more than one file in att, error.
'Er if any, export and return ToFfn. @
Const CSub$ = CMod & "ExpAtt"
Dim N%: N = AttFilCnt(D, Att)
If N <> 1 Then
    Thw CSub, "AttNm should have only one file, no export.", _
        "AttNm FilCnt ExpToFile D", _
        Att, N, ToFfn, Dbn(D)
End If
ExpAtt = ExpAttzAttd(Attd(D, Att), ToFfn)
Inf CSub, "Att is exported", "Att ToFfn FmDb", Att, ToFfn, Dbn(D)
End Function

Function ExpAttzFn$(D As Database, Att$, AttFn$, ToFfn$)
Const CSub$ = CMod & "ExpAttzFn"
If Ext(AttFn) <> Ext(ToFfn) Then
    Thw CSub, "AttFn & ToFfn are dif extEnsion|" & _
        "To export an AttFn to ToFfn, their file extEnsion should be same", _
        "AttFn-Ext ToFfn-Ext D AttNm AttFn ToFfn", _
        Ext(AttFn), Ext(ToFfn), Dbn(D), Att, AttFn, ToFfn
End If
If HasFfn(ToFfn) Then
    Thw CSub, "ToFfn Has, no over write", _
        "D AttNm AttFn ToFfn", _
        Dbn(D), Att, AttFn, ToFfn
End If
Dim Fd2 As Dao.Field2
    Set Fd2 = AttFd2(D, Att, AttFn$)

If IsNothing(Fd2) Then
    Thw CSub, "In record of AttNm there is no given AttFn, but only Act-AttFnAy", _
        "D Given-AttNm Given-AttFn Act-AttFny ToFfn", _
        Dbn(D), Att, AttFn, AttFnAy(D, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
ExpAttzFn = ToFfn
End Function

Private Function AttFd2(D As Database, Att$, AttFn$) As Dao.Field2
With Attd(D, Att)
    With .AttRs
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

Private Sub Z_ExpAtt()
Dim T$, D As Database
T = TmpFx
ExpAttzFn D, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert HasFfn(T)
Kill T
End Sub

Function FfnzFstAtt$(D As Database, Att$)
FfnzFstAtt = MovFst(Attd(D, Att).AttRs)!Filename
End Function

Function FnyzTblAtt(D As Database) As String()
FnyzTblAtt = Fny(D, "Att")
End Function

Function DAtt(D As Database) As Drs

End Function

Sub EnsTblAtt(D As Database)

End Sub


Function FnyzFldAtt(D As Database) As String()
Dim T As Dao.Recordset2: Set T = D.TableDefs("Att").OpenRecordset
Dim A As Dao.Recordset2: Set A = T!Att.Value
FnyzFldAtt = Itn(A.Fields)
End Function

Function IsOldAtt(D As Database, Att$, Ffn$) As Boolean
Const CSub$ = CMod & "IsOldAtt"
Dim TAtt As Date, TFfn As Date, AttIs$
TAtt = TimzAtt(D, Att)
TFfn = DtezFfn(Ffn)
AttIs = IIf(TAtt > TFfn, "new", "old")
Dim M$
M = "Att is " & AttIs
Inf CSub, M, "Att Ffn TimzAtt DtezFfn AttIs-Old-or-New?", Att, Ffn, TAtt, TFfn, AttIs
End Function

Function SizAtt&(D As Database, Att$)
SizAtt = ValzSsk(D, "Att", "FilSz", Av(Att))
End Function

Function TimzAtt(D As Database, Att$) As Date
TimzAtt = ValzSsk(D, "Att", "FilTim", Av(Att))
End Function

Function AttFilCntzAttd%(D As Attd)
AttFilCntzAttd = NReczRs(D.AttRs)
End Function
Function AttFilCnt%(D As Database, Att$)
AttFilCnt = AttFilCntzAttd(Attd(D, Att))
End Function

Function AttFnAy(D As Database, Att$) As String()
Dim R As Attd: R = Attd(D, Att)
AttFnAy = SyzRs(R.AttRs, "FileName")
End Function

Function AttFn$(D As Database, Att$)
AttFn = AttFnzAttd(Attd(D, Att))
End Function

Function HasOneFilAtt(D As Database, Att$) As Boolean
Debug.Print "DbAttHasOnlyFile: " & Attd(D, Att).AttRs.RecordCount
HasOneFilAtt = Attd(D, Att).AttRs.RecordCount = 1
End Function

Function AttNy(D As Database) As String()
AttNy = SyzRs(Rs(D, "Select AttNm from Att order by AttNm"))
End Function

Private Sub Z_AttFnAy()
D AttFnAy(SampDbzShpCst, "AA")
End Sub

Private Sub ZZ()
Z_AttFnAy
MDao_Att_Inf:
End Sub

Function AttNm$(A As Attd)
AttNm = A.TblRs!AttNm
End Function

Function AttFnzAttd$(A As Attd)
Const CSub$ = CMod & "AttFnzAttd"
With A.AttRs
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

Function Attd(D As Database, Att$) As Attd
With Attd
    Set .TblRs = D.OpenRecordset(FmtQQ("Select Att,FilTim,FilSz from Att where AttNm='?'", Att))
    If .TblRs.EOF Then
        D.Execute FmtQQ("Insert into Att (AttNm) values('?')", Att)
        Set .TblRs = D.OpenRecordset(FmtQQ("Select Att from Att where AttNm='?'", Att))
    End If
    Set .AttRs = .TblRs.Fields(0).Value
End With
End Function


Sub DltAtt(D As Database, Att$)
D.Execute FmtQQ("Delete * from Att where AttNm='?'", Att)
End Sub

Private Sub ImpAttzAttd(A As Attd, Ffn$)
Const CSub$ = CMod & "ImpAttzAttd"
Dim F2 As Field2
Dim S&, T As Date
S = SizFfn(Ffn)
T = DtezFfn(Ffn)
'Msg CSub, "[Att] is going to import [Ffn] with [Si] and [Tim]", FdVal(A.TblRs!AttNm), Ffn, S, T
With A
    .TblRs.Edit
    With .AttRs
        If HasReczFEv(A.AttRs, "FileName", Fn(Ffn)) Then
            D "Ffn is found in Att and it is replaced"
            .Edit
        Else
            D "Ffn is not found in Att and it is imported"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile Ffn
        .Update
    End With
    .TblRs.Fields!FilTim = DtezFfn(Ffn)
    .TblRs.Fields!FilSz = SizFfn(Ffn)
    .TblRs.Update
End With
End Sub

Sub ImpAtt(D As Database, Att$, FmFfn$)
ImpAttzAttd Attd(D, Att), FmFfn
End Sub

Private Sub Z_ImpAtt()
Dim T$, D As Database
T = TmpFt
WrtStr "sdfdf", T
ImpAtt D, "AA", T
Kill T
'T = TmpFt
'ExpAttToFfn "AA", T
'BrwFt T
End Sub

Private Sub ZZZ()
QDao_Att:
End Sub
