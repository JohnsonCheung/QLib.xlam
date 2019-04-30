Attribute VB_Name = "MVb_Dic_CntDic"
Option Explicit
Function FmtCntDic(Ay, Optional Opt As eCntOpt) As String()
FmtCntDic = FmtS1S2s(SwapS1S2s(S1S2szDic(CntDic(Ay, Opt))), Nm1:="Cnt", Nm2:="Mth")
End Function

