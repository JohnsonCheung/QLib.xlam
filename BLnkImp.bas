Attribute VB_Name = "BLnkImp"
Option Explicit
Type LiIFld
    FldNm       As String
    ShtTyLis    As String
    Extn        As String   'Extn
End Type
Type LiIFlds: N As Integer: Ay() As LiIFld: End Type

Type LiITbl
    T   As String
    Flds As LiIFlds
End Type
Type LiITbls: N As Integer: Ay() As LiITbl: End Type

Type LiIFb
    Fbn  As String
    Tbls As LiITbls
End Type
Type LiIFx
    Fxn     As String
    Tbls    As LiITbls
End Type
Type LiIFbs:  N As Integer: Ay() As LiIFb:  End Type
Type LiIFxs:  N As Integer: Ay() As LiIFx:  End Type
Type LnkImpPm
    IFxs  As LiIFxs
    IFbs  As LiIFbs
End Type
Function NewLnkImpPm(A As LiIFbs, B As LiIFxs) As LnkImpPm
With NewLnkImpPm
End With
End Function

Private Property Get Fbs() As LiIFbs
Dim L, Fbn$, T$, FF$, Fset As Aset, Bexpr$
For Each L In Itr(AywRmvT1(A, "FbTbl"))
    AsgN2tRst L, T, FF, Bexpr
    PushObj Fb, LiFb(Fbn, T, NsetzNN(FF), Bexpr)
Next
End Property

Private Property Get Fxs() As LiIFxs
Dim L, Fxn$, Wsn$, T$, Bexpr
For Each L In Itr(AywRmvT1(A, "WszT"))
    AsgN3tRst L, Fxn, Wsn, T, Bexpr
    PushObj Fx, LiFx(Fxn, Wsn, T, FxcAy(T), Bexpr)
Next
End Property

Sub LnkImp(InpFilSrc$(), LnkSrc$())
ThwIfEr ErzInpFilSrc(InpFilSrc)
ThwIfEr ErzLnkSrc(LnkSrc)
ThwIfEr ErOf_InpFil_and_LnkPm(InpFilzSrc(InpFilSrc), LnkPmzSrc(LnkInpSrc))
CrtWFb
LnkImp_Lnk
WOpn A.Apn
LnkTblzLtPm W, LtPm(A)
RunSqy W, ImpSqyzLi(A)

ThwEr ErzLidPmzV1(A), CSub
WIniOpn A.Apn
ThwEr ErzLnkTblzLtPm(W, LtPmzLid(A)), CSub
RunSqy W, ImpSqyzLidPm(A)
WCls

End Sub

Private Function ImpSqy(A As LiIFbs, B As LiIFxs) As String()
ImpSqy = SyAdd(ImpSqyzFb(A), ImpSqyzFx(B))
End Function

Private Function ImpSqyzFxs(A As LiIFxs) As String()
Dim J%, Ay() As LiIFx
Ay = A.Ay
For J = 0 To A.N - 1
    PushIAy ImpSqyzFxs, ImpSqyzFx(Ay(J))
Next
End Function

Private Function ImpSqyzFx(A As LiIFx) As String()
Dim Fm$: Fm = ">" & A.T
Dim Into$: Into = "#I" & A.T
Dim Bexpr$: Bexpr = A.Bexpr
Dim Fny$(): Fny = FnyzLidFxcAy(A.Fxc)
Dim ExtNy$(): ExtNy = ExTnyzLidFxcAy(A.Fxc)
Dim O$()
PushI O, SqlSel_FF_ExtNy_Into_Fm(Fny, ExtNy, Into, Fm, Bexpr)
ImpSqyzFx = O
End Function

Private Function ImpSqyzFb(A As LiIFb) As String()
With A
Dim FF$(): FF = .Fset.Sy
Dim Fm$: Fm = ">" & .T
Dim Into$: Into = "#I" & .T
Dim Bexpr$: Bexpr = .Bexpr
End With
Dim O$()
PushI O, "Drop table [" & Into & "]"
PushI O, SqlSel_FF_Into_Fm(FF, Into, FF, Bexpr)
ImpSqyzFb = O
End Function

Private Property Get TaxAlertLnkPmLy() As String()
Erase XX
X "@Tbl  GLBal Uom CurRate MB52 Permit PermitD SkuRepackMulti SkuTaxBy3rdParty SkuNoLongerTax"
X "@Stru GLBal Uom CurRate MB52 Permit PermitD SkuRepackMulti SkuTaxBy3rdParty SkuNoLongerTax"
X "Tbl.Where MB52 Plant='8601' and [Storage Location] in ('0002','')"
X "Tbl.Where Uom Plant='8601'"
X "Stru.Fld ZHT0.Sku       Txt Material      "
X "Stru.Fld ZHT0.CurRateAc Dbl [     Amount] "
X "Stru.Fld ZHT0.VdtFm     Txt [Valid From]  "
X "Stru.Fld ZHT0 VdtTo     Txt [Valid to] "
X "Stru.Fld ZHT0.HKD       Txt Unit"
X "Stru.Fld ZHT0.Per       Txt per "
X "Stru.Fld ZHT0.CA_Uom    Txt Uom"
X "Stru.Fld MB52.Sku    Txt Material "
X "Stru.Fld MB52.Whs    Txt Plant    "
X "Stru.Fld MB52.Loc    Txt [Storage Location] "
X "Stru.Fld MB52.BchNo  Txt Batch "
X "Stru.Fld MB52.QInsp  Dbl [In Quality Insp#]"
X "Stru.Fld MB52.QUnRes Dbl UnRestricted"
X "Stru.Fld MB52.QBlk   Dbl Blocked"
X "Stru.Fld MB52.VInsp  Dbl [Value in QualInsp#] "
X "Stru.Fld MB52.VUnRes Dbl [Value Unrestricted] "
X "Stru.Fld MB52.VBlk   Dbl [Value BlockedStock]"
X "Stru.Fld Uom.Sku     Txt Material "
X "Stru.Fld Uom.Des     Txt [Material Description] "
X "Stru.Fld Uom.AC_U    Txt [Unit per case] "
X "Stru.Fld Uom.SkuUom  Txt [Base Unit of Measure] "
X "Stru.Fld Uom.BusArea Txt [Business Area]"
X "Stru.Fld GLBal.BusArea Txt [Business Area Code]"
X "Stru.Fld GLBal.GLBal   Dbl                     "
X "Stru.Fld Permit           GLBal   Dbl                     "
X "Stru.Fld PermitD          GLBal   Dbl                     "
X "Stru.Fld SkuRepackMulti   GLBal   Dbl                     "
X "Stru.Fld SkuTaxBy3rdParty GLBal   Dbl                     "
X "Stru.Fld  SkuNoLongerTax"
X "#     T    Fx   Ws   S"
X "FxTbl Z86  ZHT1.8601 ZHT1"
X "FxTbl Z87  ZHT1.8701 ZHT1"
X "FxTbl Uom  ZHT1      Uom"
X "FxTbl Uom  ZHT1      Uom"
X "#   Stru.Fld    S Extn       # S = ShtTy"

X "Fld ZHT1.ZHT1   D Brand  "
X "Fld ZHT1.RateSc M Amount "
X "Fld ZHT1.VdtFm  M [Valid From]  "
X "Fld ZHT1.VdtTo  M [Valid to]"
X "Fld Uom.Sku    M Material "
X "Fld Uom.Des    M [Material Description] "
X "Fld Uom.Sc_U   M SC "
X "Fld Uom.StkUom M [Base Unit of Measure] "
X "Fld Uom.Topaz  M [Topaz Code] "
X "Fld Uom.ProdH  M [Product hierarchy]"
X "Fld MB52.Sku    M Material "
X "Fld MB52.Whs    M Plant    "
X "Fld MB52.QInsp  D [In Quality Insp#]"
X "Fld MB52.QUnRes D Unrestricted"
X "Fld MB52.QBlk   D Blocked"
SampLnkPmLy = XX
Erase XX
End Property


Function StruDotFld(T$) As String()

Dim A$() 'Each Lin is [T.F ShtTy]
StruDotFld = SyAddPfx(SyAlignNCol(A), "Stru.Fld ")
End Function
Function StruDotFldzFxw(Fx$, Wsn$) As String()

End Function


Private Sub Z_LnkImp()
Dim Db As Database
Set Db = LnkImp(SampLnkPm)
BrwFb A.Name
Stop
End Sub

Private Sub Z_ChkColzLnkPm()
Brw ChkColzLnkPm(ShpCstLnkPm)
End Sub

Private Function ImpSqlzFx$(A As LiIFx)
Dim Fm$: Fm = ">" & A.T
Dim Into$: Into = "#I" & A.T
Dim Bexpr$: Bexpr = A.Bexpr
ImpSqlzFx = SqlSel_FF_ExtNy_Into_Fm(A.Fny, A.ExtNy, Into, Fm, Bexpr)
End Function

Private Function ImpSqlzFb$(A As LiIFb)
With A
Dim FF$(): FF = .Fset.Sy
Dim Fm$: Fm = ">" & .T
Dim Into$: Into = "#I" & .T
Dim Bexpr$: Bexpr = .Bexpr
End With
ImpSqlzFb = SqlSel_FF_Into_Fm(FF, Into, FF, Bexpr)
End Function

Property Get SampLnkImpPm() As LnkImpPm
Set SampLnkPm = LnkPm(SampLnkPmSrc)
End Property

Function LnkImpPm(Src$()) As LnkImpPm
If Si(Src) = 0 Then Thw CSub, "No lines in Src"
If T1(Src(0)) <> "LidPm" Then Thw CSub, "First line must be LidPm", "Src", Src
LnkPmzSrc = LnkPm()
With LnkPm
    
End With
LidPm.Init Apn, Fil, Fx, Fb
End Function

Function LnkImpSrcEr(LnkImpSrc$()) As String()
End Function

Function SampSrczLnkImp() As String()
Erase XX
X "Nm ShpCst"
X "WszT ZHT1 8701   ZHT18701"
X "WszT ZHT1 8601   ZHT18601"
X "WszT MB52 Sheet1 MB52"
X "WszT UOM  Sheet1 UOM"
X "WsCol ZHT18701 ZHT1   M Brand"
X "WsCol ZHT18701 RateSc D Amount"
X "WsCol ZHT18701 VdtFm  M Valid From"
X "WsCol ZHT18701 VdtTo  M Valid to"
X "WsCol ZHT18601 ZHT1   M Brand"
X "WsCol ZHT18601 RateSc D Amount"
X "WsCol ZHT18601 VdtFm  M Valid From"
X "WsCol ZHT18601 VdtTo  M Valid to"
X "WsCol UOM Sku     M Material"
X "WsCol UOM Des     M Material Description"
X "WsCol UOM Sc_U    M SC "
X "WsCol UOM StkUom  M Base Unit of Measure"
X "WsCol UOM Topaz   M Topaz Code"
X "WsCol UOM ProdH   M Product hierarchy"
X "WsCol MB52 Sku    M Material"
X "WsCol MB52 Whs    M Plant"
X "WsCol MB52 QInsp  D In Quality Insp#"
X "WsCol MB52 QUnRes D Unrestricted"
X "WsCol MB52 QBlk   D Blocked"
SampLnkPmSrc = XX
Erase XX
End Function


Private Function LnkTblPmszFx() As LnkTblPms
'Private Function LnkTblPmszFx(A As LiFxs, FfnDic As Dictionary) As LnkTblPms
Dim J%, Fx$, M As LtPm
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        'Fx = FfnDic(.Fxn)
        PushLnkTblPm O, NewLnkTblPm()
        'LtPmAyFx, M.Init(">" & .T, .Wsn & "$", CnStrzFxDAO(Fx$))
    End With
Next
End Function

Private Function LnkPmszFbs(A As LiIFbs, FfnDic As Dictionary) As LnkTblPms
Dim J%, Fb$, Ay() As LiIFb, O As LnkLnkPms
For J = 0 To A.N - 1
    With Ay(J)

        Fb = FfnDic(.Fbn)
        S = 1
        T = ">" & 1
        Cn = LtPmAyFb , M.Init(">" & .T, .T, CnStrzFxAdo(Fb$))

        PushLnkTblPm O, NewLnkTblPm(T, S, Cn)
    End With
Next
End Function




