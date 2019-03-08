Attribute VB_Name = "MDao_Li_Pm"
Option Explicit
Private A$()
Sub BrwSampLiPm()
SampLiPm.Brw
End Sub
Property Get SampLiPm() As LiPm
Set SampLiPm = LiPm(SampLiPmSrc)
End Property
Function LiPm(Src$()) As LiPm
A = Src
Set LiPm = New LiPm
LiPm.Init Apn, Fil, Fx, Fb
End Function

Private Property Get Apn$()
Dim X$: X = A(0)
Apn = T2(X)
If T1(X) <> "LiPm" Then Thw CSub, "T1 of first line of src must be LiPm", "First-Line-of-LiPm-Src Src", A(0), A
End Property

Private Property Get Fil() As LiFil()
Dim O$(): O = AywRmvT1(A, "Fil")
If Sz(O) = 0 Then Thw CSub, "There is no Fil-Line in LiPmSrc", "LiPmSrc", A
Dim L, FilNm$, Ffn$
For Each L In O
    AsgTRst L, FilNm, Ffn
    PushObj Fil, LiFil(FilNm, Ffn)
Next
End Property

Private Property Get Fb() As LiFb()
Dim L, Fbn$, T$, FF$, Fset As Aset, Bexpr$
For Each L In Itr(AywRmvT1(A, "FbTbl"))
    Asg2TRst L, T, FF, Bexpr
    PushObj Fb, LiFb(Fbn, T, FsetFF(FF), Bexpr)
Next
End Property

Private Property Get Fx() As LiFx()
Dim L, Fxn$, Wsn$, T$, Bexpr
For Each L In Itr(AywRmvT1(A, "WszT"))
    Asg3TRst L, Fxn, Wsn, T, Bexpr
    PushObj Fx, LiFx(Fxn, Wsn, T, FxcAy(T), Bexpr)
Next
End Property

Private Function FxcAy(T$) As LiFxc()
FxcAy = LiFxcAyLnkColAy(AywRmvTT(A, "WsCol", T))
End Function
Function SampLiPmSrc() As String()
Erase XX
X "LiPm ShpCst"
X "Fil MB52 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\MB52 2018-07-30.xls"
X "Fil UOM  C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\sales text.xlsx"
X "Fil ZHT1 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\ZHT1.XLSX"
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
SampLiPmSrc = XX
Erase XX
End Function
Private Function LtPmAyFx(A() As LiFx, FfnDic As Dictionary) As LtPm()
Dim J%, Fx$, M As LtPm
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        Fx = FfnDic(.Fxn)
        PushObj LtPmAyFx, M.Init(">" & .T, .Wsn & "$", CnStrzFxDAO(Fx))
    End With
Next
End Function
Private Function LtPmAyFb(A() As LiFb, FfnDic As Dictionary) As LtPm()
Dim J%, Fb$, M As LtPm
For J = 0 To UB(A)
    Set M = New LtPm
    With A(J)
        Fb = FfnDic(.Fbn)
        PushObj LtPmAyFb, M.Init(">" & .T, .T, CnStrzFxAdo(Fb))
    End With
Next
End Function

Function LtPm(A As LiPm) As LtPm()
Dim O() As LtPm, D As Dictionary
Set D = A.FilNmToFfnDic
PushObjAy O, LtPmAyFb(A.Fb, D)
PushObjAy O, LtPmAyFx(A.Fx, D)
LtPm = O
End Function


