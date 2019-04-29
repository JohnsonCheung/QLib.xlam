Attribute VB_Name = "MVb_Dic_CntDic"
Option Explicit
Function FmtCntDic(Ay, Optional Opt As eCntOpt) As String()
FmtCntDic = FmtS1S2Ay(SwapS1S2Ay(S1S2AyzDic(CntDic(Ay, Opt))), Nm1:="Cnt", Nm2:="Mth")
End Function

