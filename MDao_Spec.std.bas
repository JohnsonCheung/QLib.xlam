Attribute VB_Name = "MDao_Spec"
Option Explicit
Const CMod$ = "MDao_Spec."

Sub DbSPnmImp(A As Database, Spnm)
Const CSub$ = CMod & "DbSPnmImp"
Dim Ft$
'    Ft = SpnmFt(Spnm)
    
Dim NoCur As Boolean
Dim NoLas As Boolean
Dim CurOld As Boolean
Dim CurNew As Boolean
Dim SamTim As Boolean
Dim DifSz As Boolean
Dim SamSz As Boolean
Dim DifFt As Boolean
Dim Rs As Dao.Recordset
    Q = FmtQQ("Select SpecNm,Ft,Lines,Tim,Sz,LTimStr_Dte from Spec where SpecNm = '?'", Spnm)
    Set Rs = CDb.OpenRecordset(Q)
    NoCur = Not HasFfn(Ft)
    'NoLas = HasRec(Rs)
    
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    Dim CurS&, LasS&
    Dim LasFt$, LdTimStr_Dte$
    CurS = FfnSz(Ft)
    CurT = TimFfn(Ft)
    If Not NoLas Then
        With Rs
            LasS = Nz(Rs!Sz, -1)
            LasT = Nz(!Tim, 0)
            LasFt = Nz(!Ft, "")
'            LdTimStr_Dte = DteTimStr(!LTimStr_Dte)
        End With
    End If
    SamTim = CurT = LasT
    CurOld = CurT < LasT
    CurNew = CurT > LasT
    SamSz = CurS = LasS
    DifSz = Not SamSz
    DifFt = Ft <> LasFt
    

Const Imported$ = "***** IMPORTED ******"
Const NoImport$ = "----- no import -----"
Const NoCur______$ = "No Ft."
Const NoLas______$ = "No Last."
Const FtDif______$ = "Ft is dif."
Const SamTimSz___$ = "Sam tim & sz."
Const SamTimDifSz$ = "Sam tim & sz. (Odd!)"
Const CurIsOld___$ = "Cur is old."
Const CurIsNew__$ = "Cur is new."
Const C$ = "|[SpecNm] [Db] [Cur-Ft] [Las-Ft] [Cur-Tim] [Las-Tim] [Cur-Sz] [Las-Sz] [Imported-Time]."

Dim Dr()
Dr = Array(Spnm, Ft, FtLines(Ft), CurT, CurS, Now)
Select Case True
Case NoCur, SamTim:
'Case NoLas: InsDrzRs Dr, Rs
'Case DifFt, CurNew: Dr_Upd_Rs Dr, Rs
Case Else: Stop
End Select

Dim Av()
Av = Array(Spnm, DbNm(A), Ft, LasFt, CurT, LasT, CurS, LasS, LdTimStr_Dte)
Select Case True
'Case NoCur:            XDmp_Lin_AV CSub, NoImport & NoCur______ & C, Av
'Case NoLas:            XDmp_Lin_AV CSub, Imported & NoLas______ & C, Av
'Case DifFt:            XDmp_Lin_AV CSub, Imported & FtDif______ & C, Av
'Case SamTim And SamSz: XDmp_Lin_AV CSub, NoImport & SamTimSz___ & C, Av
'Case SamTim And DifSz: XDmp_Lin_AV CSub, NoImport & SamTimDifSz & C, Av
'Case CurOld:           XDmp_Lin_AV CSub, NoImport & CurIsOld___ & C, Av
'Case CurNew:           XDmp_Lin_AV CSub, Imported & CurIsNew__ & C, Av
Case Else: Stop
End Select
End Sub

Property Get SpecPth$(Apn$)
SpecPth = PthEns(WPth(Apn) & "Spec\")
End Property

Sub BrwSpecPth(Apn$)
BrwPth SpecPth(Apn)
End Sub

Sub ClrSpecPth(Apn$)
ClrPth SpecPth(Apn)
End Sub

Property Get SpecSchm() As String()
SpecSchm = SplitVBar(SpecSchmVbl)
End Property

Sub EnsSpecTblz(A As Database)
'If Not HasTbl(A, "Spec") Then CrtSpecTbl A
End Sub

Sub CrtSpecTblz(A As Database)
CrtSchm A, SpecSchm
End Sub

Sub EnsSpecTbl()
EnsSpecTblz CDb
End Sub

Sub ExpSpecz(A As Database)
'ClrSpecPth A
Dim N
For Each N In Itr(SpecNy)
'    ExpSpnm N
Next
End Sub

Property Get SpecNy() As String()
SpecNy = SpecNyDb(CDb)
End Property

Function SpecNyDb(A As Dao.Database) As String()
'SpecNyDb = SyDbtf(A, "Spec", "SpecNm")
End Function


