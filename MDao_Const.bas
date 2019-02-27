Attribute VB_Name = "MDao_Const"
Option Explicit
Public Const StdEleLines$ = _
"E Crt Dte;Req;Dft=Now" & vbCrLf & _
"E Tim Dte" & vbCrLf & _
"E Lng Lng" & vbCrLf & _
"E Mem Mem" & vbCrLf & _
"E Dte Dte" & vbCrLf & _
"E Nm  Txt;Req;Sz=50"
Public Const StdETFLines$ = _
"ETF Nm  * *Nm          " & vbCrLf & _
"ETF Tim * *Tim         " & vbCrLf & _
"ETF Dte * *Dte         " & vbCrLf & _
"ETF Crt * CrtTim       " & vbCrLf & _
"ETF Lng * Sz           " & vbCrLf & _
"ETF Mem * Lines *Ft *Fx"

Public Const SpecSchmVbl$ = "Tbl Spec * SpecNm | Lines Ft Sz Tim LTimStr CrtTim"
Public Const C_Des$ = "Description"

Public Const SampReSeqSpec$ = _
"Flg RecTy Amt Key Uom MovTy Qty BchRateUX RateTy Bch Las GL |" & _
" Flg IsAlert IsWithSku |" & _
" Key Sku PstMth PstDte |" & _
" Bch BchNo BchPermitDate BchPermit |" & _
" Las LasBchNo LasPermitDate LasPermit |" & _
" GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
" Uom Des StkUom Ac_U"
Public Const ShtTyzDaosl$ = "Byt Mem Lng Int Dte Txt Yes Dbl Sng Cur"
Public Const LgSchmNm$ = "LgSchm" ' The LgSchm-Spnm

Private Property Get LnkSpecTp$()
Const A_1$ = "E Mem | Mem Req AlwZLen" & _
vbCrLf & "E Txt | Txt Req" & _
vbCrLf & "E Crt | Dte Req Dft=Now" & _
vbCrLf & "E Dte | Dte" & _
vbCrLf & "F Amt * | *Amt" & _
vbCrLf & "F Crt * | CrtDte" & _
vbCrLf & "F Dte * | *Dte" & _
vbCrLf & "F Txt * | Fun * Txt" & _
vbCrLf & "F Mem * | Lines" & _
vbCrLf & "T Sess | * CrtDte" & _
vbCrLf & "T Msg  | * Fun *Txt | CrtDte" & _
vbCrLf & "T Lg   | * Sess Msg CrtDte" & _
vbCrLf & "T LgV  | * Lg Lines" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt" & _
vbCrLf & "D . Msg | ..."

LnkSpecTp = A_1
End Property
