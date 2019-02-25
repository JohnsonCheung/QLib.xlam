Attribute VB_Name = "MVb_Fs_Ffn_Kind"
Option Explicit

Function TblKd$(Ffn)
Select Case True
Case IsFx(Ffn): TblKd = "worksheet"
Case IsFb(Ffn): TblKd = "table"
Case Else: TblKd = "[TblKd=unknown]"
End Select
End Function
Function FfnKd$(Ffn)
Select Case True
Case IsFx(Ffn): FfnKd = "access-file"
Case IsFb(Ffn): FfnKd = "excel-file"
Case Else: FfnKd = "[FfnKd=unknown]"
End Select
End Function
