Attribute VB_Name = "MxSpec"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxSpec."
Public Const PrimNN$ = "Boolean Byte Integer Int Long Single Double Char Text Memo Attachment" ' used in TzPFld

Sub AddFdy(A As TableDef, Fdy() As dao.Field2)
Dim I: For Each I In Fdy
    A.Fields.Append I
Next
End Sub

Sub BrwSpecPth(Apn$)
BrwPth SpecPth(Apn)
End Sub

Sub CrtTblSpec(D As Database)
CrtSchm D, SplitVBar(SpecSchmVbl)
End Sub

Sub CrtTblzPFld(D As Database, T, PrimNN$)
'Fm PrimNN: #Sql-FldLis-Phrase.  !The fld spec of create table sql inside the bkt.  Each fld sep by comma.  The spec allows:
'                                 !Boolean Byte Integer Int Long Single Double Char Text Memo Attachment
'Ret : create the @T in @D by DAO @@
Dim Td As dao.TableDef: Set Td = TdzNm(T)
AddFdy Td, FdAy(PrimNN)
D.TableDefs.Append Td
'AddFdy D.TableDefs(T), FdAy(PrimNN)
End Sub

Sub EnsTblSpec(D As Database)
If Not HasTbl(D, "Spec") Then CrtTblSpec D
End Sub

Sub ExpSpec(D As Database, ToPth$)
ClrPth ToPth
Dim N$, I
For Each I In Itr(SpecNy(D))
    N = I
    ExpSpeczNm D, N, ToPth
Next
End Sub

Sub ExpSpeczNm(D As Database, SpecNm$, ToPth$)

End Sub

Function FdAy(Prim$) As dao.Field2()
'Fm PrimNN: #Sql-FldLis-Phrase.  !The fld spec of create table sql inside the bkt.  It allows attachment.  It uses DAO to create
Dim F: For Each F In Itr(AmTrim(SplitComma(PrimNN)))
    PushObj FdAy, FdzPFld(F)
Next
End Function

Function FdzPFld(PFld) As dao.Field2
'Fm PFld: #Sql-Fld-Phrase.  !The single fld spec of create table sql inside the bkt.  It allows attachment.
Dim N$, S$ ' #Fldn and #Spec
Dim O As dao.Field2
AsgBrkSpc PFld, N, S
Select Case True
Case S = "Boolean":  Set O = FdzBool(N)
Case S = "Byte":     Set O = FdzByt(N)
Case S = "Integer", S = "Int": Set O = FdzInt(N)
Case S = "Long":     Set O = FdzLng(N)
Case S = "Single":   Set O = FdzSng(N)
Case S = "Double":   Set O = FdzDbl(N)
Case S = "Currency": Set O = FdzCur(N)
Case S = "Char":     Set O = FdzChr(N)
Case HasPfx(S, "Text"): Set O = FdzTxt(N, BetBkt(S))
Case S = "Memo":     Set O = FdzMem(N)
Case S = "Attachment": Set O = FdzAtt(N)
Case S = "Time":     Set O = FdzTim(N)
Case S = "Date":     Set O = FdzDte(N)
Case Else: Thw CSub, "Invalid PFld", "Nm Spec vdt-PFld", N, S, PrimNN
End Select
Set FdzPFld = O
End Function

Sub ImpSpec(D As Database, Spnm)
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
Dim Rs As dao.Recordset
    Dim Q$
    Q = FmtQQ("Select SpecNm,Ft,Lines,Tim,Si,LTimStr_Dte from Spec where SpecNm = '?'", Spnm)
    Set Rs = D.OpenRecordset(Q)
    NoCur = NoFfn(Ft)
    'NoLas = HasRec(Rs)
    
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    Dim CurS&, LasS&
    Dim LasFt$, LdTimStr_Dte$
    CurS = SizFfn(Ft)
    CurT = DtezFfn(Ft)
    If Not NoLas Then
        With Rs
            LasS = Nz(Rs!Si, -1)
            LasT = Nz(!Tim, 0)
            LasFt = Nz(!Ft, "")
'            LdTimStr_Dte = TimStr(!LTimStr_Dte)
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
Dr = Array(Spnm, Ft, LineszFt(Ft), CurT, CurS, Now)
Select Case True
Case NoCur, SamTim:
'Case NoLas: InsDrzRs Dr, Rs
'Case DifFt, CurNew: Dr_Upd_Rs Dr, Rs
Case Else: Stop
End Select

Dim Av()
Av = Array(Spnm, D.Name, Ft, LasFt, CurT, LasT, CurS, LasS, LdTimStr_Dte)
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

Function SpecNy(D As Database) As String()
SpecNy = SyzTF(D, "Spec", "SpecNm")
End Function

Function SpecPth$(WPth$)
SpecPth = EnsPth(WPth & "Spec\")
End Function
