Attribute VB_Name = "MxDcitm"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDcitm."
':Dcn: :Nm #Dcl-Itm-Nm#
':Dcitm: :S #Dcl-Itm# ! Dc
Function Dcn$(Dcitm)
If HasSubStr(Dcitm, " As ") Then
    Dcn = DcnoAs(Dcitm)
Else
    Dcn = DcnoTyChr(Dcitm)
End If
End Function

Private Function DcnoTyChr$(DimShtItm)
DcnoTyChr = RmvLasChrzzLis(RmvSfxzBkt(DimShtItm), MthTyChrLis)
End Function

Private Function DcnoAs$(DimAsItm)
DcnoAs = RmvSfxzBkt(Bef(DimAsItm, " As"))
End Function
