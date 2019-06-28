Attribute VB_Name = "QVb_Ay_Ixy"
Option Compare Text
Option Explicit
Enum EmIxFm 'Ix is always 0.  When present as Lno, Ix=0 will be Lno=1.
    EiFm1   'Or, presenting Ix=0 as 1 for Lno
    EiFm0   'So, presenting Ix=0 as 0  for Ix
End Enum

Function Lnoss$(Ixy() As Long)
Lnoss = JnSpc(AyIncEle1(Ixy))
End Function

