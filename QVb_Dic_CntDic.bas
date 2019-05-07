Attribute VB_Name = "QVb_Dic_CntDic"
Option Explicit
Private Const CMod$ = "MVb_Dic_CntDic."
Private Const Asm$ = "QVb"
Function FmtCntDic(Ay, Optional Opt As EmCnt) As String()
FmtCntDic = FmtS1S2s(SwapS1S2s(S1S2szDic(CntDic(Ay, Opt))), Nm1:="Cnt", Nm2:="Mth")
End Function

