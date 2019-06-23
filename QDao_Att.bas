Attribute VB_Name = "QDao_Att"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Att."
Private Type Attd
    TblRs As DAO.Recordset '..Att.. #Tbl-Rs ! It is the Tbl-Att recordset
    AttRs As DAO.Recordset '.       #Att-Rs2 !
End Type
'TblAtt : AttNm Att FilSi FilTim ' Inside @Att it expected to have only one att file, becuase there is only one FilSi and FilTim

Private Sub XThwIf_ExtDif(Fun$, A As Attd, ToFfn$)
With A.AttRs
    If Ext(!Filename) <> Ext(ToFfn) Then Thw Fun, "The Ext in the Att should be same", "Att-Ext ToFfn-Ext", Ext(!Filename), Ext(ToFfn)
End With
End Sub

Private Function XF2(A As Attd) As DAO.Field2
Set XF2 = A.AttRs!FileData
End Function

Private Function XAttNm$(A As Attd)
XAttNm = A.TblRs!AttNm
End Function
Private Sub XThwIf_CntNe1(Fun$, D As Database, Att$)
Dim N%: N = XNAttInFld(D, Att)
If N <> 1 Then
    Thw Fun, "AttNm should have only one file, no export.", _
        "AttNm FilCnt D", _
        Att, N, Dbn(D)
End If
End Sub
Private Function XAtt$(A As Attd)
XAtt = A.TblRs!Att
End Function
Function ExpAtt$(D As Database, Att$, ToFfn$)
'Ret Exporting the first File in [Att] to [ToFfn] if Att is newer or ToFfn not exist.
'Er if no or more than one file in att, error.
'Er if any, export and return ToFfn. @@
XThwIf_CntNe1 CSub, D, Att
Dim A As Attd: A = XAttd(D, Att)
XThwIf_ExtDif CSub, A, ToFfn
XF2(A).SaveToFile ToFfn
ExpAtt = ToFfn
Inf CSub, "Att is exported", "Att ToFfn FmDb", XAtt(A), ToFfn, Dbn(D)
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
Dim Fd2 As DAO.Field2
    Set Fd2 = XF2zFn(D, Att, AttFn$)

If IsNothing(Fd2) Then
    Thw CSub, "In record of AttNm there is no given AttFn, but only Act-AttFnAy", _
        "D Given-AttNm Given-AttFn Act-AttFny ToFfn", _
        Dbn(D), Att, AttFn, XAttFnAy(D, Att), ToFfn
End If
Fd2.SaveToFile ToFfn
ExpAttzFn = ToFfn
End Function

Private Function XF2zFn(D As Database, Att$, AttFn$) As DAO.Field2
With XAttd(D, Att)
    With .AttRs
        .MoveFirst
        While Not .EOF
            If !Filename = AttFn Then
                Set XF2zFn = !FileData
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

Function FnyzAttTbl(D As Database) As String()
FnyzAttTbl = Fny(D, "Att")
End Function

Function DAtt(D As Database) As Drs
End Function
Function DTblAtt(D As Database) As Drs
DTblAtt = DrszT(D, "Att")
End Function
Private Sub Z_EnsTblAtt()
Dim D As Database: Set D = TmpDb
EnsTblAtt D
End Sub
Function TmpAttDb() As Database
'Ret: a tmp db with tbl-att @@
Dim O As Database: Set O = TmpDb
EnsTblAtt O
Set TmpAttDb = O
End Function
Sub EnsTblAtt(D As Database)
If HasTbl(D, "Att") Then
    Dim FF$: FF = FFzT(D, "Att")
    If FF <> "AttNm Att FilSi FilTim" Then Thw CSub, "Db has :Tbl:Att, but its FF is not [AttNm Att FilSi FilTim", "Dbn Tbl-Att-FF", Dbn(D), FF
End If
Dim PFld$: PFld = "AttNm Text(255), Att Attachment, FilSi Long,FilTim Date" ' #Sql-Fld-Phrase.  The fld spec of create table sql inside the bkt.
CrtTblzPFld D, "Att", PFld
DbzReOpn(D).Execute "Create Index PrimaryKey on Att (AttNm) with Primary"
End Sub
Function DAttFld(A As Attd) As Drs
'Ret : :Drs:Fldn DtaTy Si: of the Fld-Att of Tbl-Att.  The Fld-Att is Dao.Recordset2
DAttFld = DFldzRs(A.AttRs)
End Function
Function DAttFldzDb(D As Database) As Drs
'Ret : :Drs:Fldn DtaTy Si: from @D assume there is table-Att
DAttFldzDb = DAttFld(XAttd(D, "Sample"))
End Function

Function DFldzRs(R As DAO.Recordset2) As Drs
'Ret : :Drs:Fldn DtaTy Si: the @R
Dim Dy(), F As DAO.Field2: For Each F In R.Fields
    Dim N$: N = F.Name
    Dim T$: T = DtaTy(F.Type)
    Dim S%: S = F.Size
    PushI Dy, Array(N, T, S)
Next
DFldzRs = DrszFF("Fldn DtaTy Si", Dy)
End Function

Private Sub Z_FnyzAttFld()
Dim D As Database: Set D = TmpDb
EnsTblAtt D
Dim Ft$: Ft = TmpFt
WrtStr "lskdfjldf", Ft
ImpAtt D, "Txt", Ft
Dmp FnyzAttFld(D)
End Sub
Function FnyzAttFld(D As Database) As String()
'Ret : fny of the att fld in att tbl of @D.
'      Assume there is tbl "Att" in @D
'      Assume there fld "Att" in tbl "Att"
'      Assume the fld is Dao.Recordset2
'      return the fny of such Recordset2 @@
Dim T As DAO.Recordset2: Set T = D.TableDefs("Att").OpenRecordset
Dim A As DAO.Recordset2: Set A = T!Att.Value
FnyzAttFld = Itn(A.Fields)
End Function

Function XIsAttOld(D As Database, Att$, Ffn$) As Boolean
Const CSub$ = CMod & "XIsAttOld"
Dim ATim$:   ATim = XAttTim(D, Att)
Dim FTim$:   FTim = DtezFfn(Ffn)
Dim AttIs$: AttIs = IIf(ATim > FTim, "new", "old")
Dim M$:         M = "Att is " & AttIs
Inf CSub, M, "Att Ffn AttTim FfnTim AttIs-Old-or-New?", Att, Ffn, ATim, FTim, AttIs
End Function

Function XAttSi&(D As Database, Att$)
XAttSi = ValzSsk(D, "Att", "FilSz", Av(Att))
End Function

Function XAttTim(D As Database, Att$) As Date
XAttTim = ValzSsk(D, "Att", "FilTim", Av(Att))
End Function

Function XNAttInTbl%(D As Attd)
XNAttInTbl = NReczRs(D.AttRs)
End Function
Function XNAttInFld%(D As Database, Att$)
XNAttInFld = XNAttInTbl(XAttd(D, Att))
End Function

Private Function XAttFnAy(D As Database, Att$) As String()
Dim R As Attd: R = XAttd(D, Att)
XAttFnAy = SyzRs(R.AttRs, "FileName")
End Function

Private Function XAttFn$(D As Database, Att$)
'Ret : fst attachment fn in the att fld of att tbl, if no fn, return blnk @@
Const CSub$ = CMod & "AttFnzAttd"
Dim A As Attd: A = XAttd(D, Att) ' if @Att in exist in Tbl-Att, a rec will created
With A.AttRs
    If .EOF Then
        If .BOF Then
            Inf CSub, "[AttNm] has no attachment files", "AttNm", AttNm(A)
            Exit Function
        End If
    End If
    .MoveFirst
    XAttFn = !Filename
End With
End Function

Function XIsAttOneFil(D As Database, Att$) As Boolean
Debug.Print "DbAttHasOnlyFile: " & XAttd(D, Att).AttRs.RecordCount
XIsAttOneFil = XAttd(D, Att).AttRs.RecordCount = 1
End Function

Private Function XAttNy(D As Database) As String()
XAttNy = SyzRs(Rs(D, "Select AttNm from Att order by AttNm"))
End Function

Private Sub Z_XAttFnAy()
D XAttFnAy(SampDbzShpCst, "AA")
End Sub

Private Sub Z()
QDao_Att:
End Sub

Function AttNm$(A As Attd)
AttNm = A.TblRs!AttNm
End Function

Private Function XAttd(D As Database, Att$) As Attd
'Ret: :Attd ! which keeps :TblRs and :AttRs opened,
'           ! where :TblRs is poiting the rec in tbl-att, if fnd just point to it, if not fnd, add one rec with AttNm=@Att
'           ! and   :AttRs is pointing to the :FileData of the fld-Att of the tbl-Att
Dim Q$: Q = FmtQQ("Select Att,FilTim,FilSi from Att where AttNm='?'", Att)
If Not HasReczQ(D, Q) Then
    D.Execute FmtQQ("Insert into Att (AttNm) values('?')", Att) ' add rec to tbl-att with Att=@Att
End If
With XAttd
    Set .TblRs = Rs(D, Q)
    Set .AttRs = .TblRs.Fields(0).Value ' there is always a rec of Att=@Att in .TblRs (Tbl-Att)
End With
End Function

Sub DltAtt(D As Database, Att$)
D.Execute FmtQQ("Delete * from Att where AttNm='?'", Att)
End Sub

Sub ImpAtt(D As Database, Att$, FmFfn$)
Dim F2 As DAO.Field2
'Msg CSub, "[Att] is going to import [Ffn] with [Si] and [Tim]", FdVal(A.TblRs!AttNm), Ffn, S, T
Dim A As Attd: A = XAttd(D, Att)
Dim T As DAO.Recordset2: Set T = A.TblRs ' The Tbl-Rs of Tbl-Att
    T.Edit
    With A.AttRs
        If HasReczFEv(A.AttRs, "FileName", Fn(FmFfn)) Then
            Dmp "Ffn is found in Att and it is replaced"
            .Edit
        Else
            Dmp "Ffn is not found in Att tbl and it is IMPORTED.  Ffn[" & FmFfn & "]"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile FmFfn
        .Update
    End With
    A.TblRs.Fields!FilTim = DtezFfn(FmFfn)
    A.TblRs.Fields!FilSi = SizFfn(FmFfn)
    A.TblRs.Update
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

