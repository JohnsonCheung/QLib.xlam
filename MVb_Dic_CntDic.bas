Attribute VB_Name = "MVb_Dic_CntDic"
Option Explicit
Function FmtCntDic(Ay, Optional IgnCas As Boolean, Optional Opt As eCntOpt) As String()
FmtCntDic = FmtS1S2Ay(SwapS1S2Ay(S1S2AyzDic(CntDic(Ay, IgnCas, Opt))), Nm1:="Cnt", Nm2:="Mth")
End Function

