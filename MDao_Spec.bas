Attribute VB_Name = "MDao_Spec"
Option Explicit
Const CMod$ = "MDao_Spec."

Sub ImpSpec(A As Database, Spnm)
Const CSub$ = CMod & "ImpSpec"
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
    Q = FmtQQ("Select SpecNm,Ft,Lines,Tim,Si,LTimStr_Dte from Spec where SpecNm = '?'", Spnm)
    Set Rs = A.OpenRecordset(Q)
    NoCur = Not HasFfn(Ft)
    'NoLas = HasRec(Rs)
    
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    Dim CurS&, LasS&
    Dim LasFt$, LdTimStr_Dte$
    CurS = FfnSz(Ft)
    CurT = TimFfn(Ft)
    If Not NoLas Then
        With Rs
            LasS = Nz(Rs!Si, -1)
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
Const SamTimSi___$ = "Sam tim & sz."
Const SamTimDifSz$ = "Sam tim & sz. (Odd!)"
Const CurIsOld___$ = "Cur is old."
Const CurIsNew__$ = "Cur is new."
Const C$ = "|[SpecNm] [Db] [Cur-Ft] [Las-Ft] [Cur-Tim] [Las-Tim] [Cur-Si] [Las-Si] [Imported-Time]."

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
'Case SamTim And SamSz: XDmp_Lin_AV CSub, NoImport & SamTimSi___ & C, Av
'Case SamTim And DifSz: XDmp_Lin_AV CSub, NoImport & SamTimDifSz & C, Av
'Case CurOld:           XDmp_Lin_AV CSub, NoImport & CurIsOld___ & C, Av
'Case CurNew:           XDmp_Lin_AV CSub, Imported & CurIsNew__ & C, Av
Case Else: Stop
End Select
End Sub

Function SpecPth$(Apn$)
SpecPth = PthEns(WPth(Apn) & "Spec\")
End Function

Sub BrwSpecPth(Apn$)
BrwPth SpecPth(Apn)
End Sub

Sub ClrSpecPth(Apn$)
ClrPth SpecPth(Apn)
End Sub

Sub EnsTblSpec(A As Database)
'If Not HasTbl(A, "Spec") Then CrtSpecTbl A
End Sub

Sub CrtTblSpec(A As Database)
CrtSchm A, SplitVBar(SpecSchmVbl)
End Sub

Sub ExpSpec(Apn$)
ClrSpecPth Apn
Dim N
For Each N In Itr(SpecNy(Apn))
    'ExpSpnm N
Next
End Sub

Function SpecNy(Apn$) As String()
SpecNy = SyzTF(AppDb(Apn), "Spec", "SpecNm")
End Function


