Attribute VB_Name = "MVb_Str_UnderLin"
Option Explicit
Function UnderLin$(Lin$)
UnderLin = String(Len(Lin), "-")
End Function

Function UnderLinDbl$(Lin$)
UnderLinDbl = String(Len(Lin), "=")
End Function

Function PushMsgUnderLinDbl(O$(), M$)
Push O, M
Push O, UnderLinDbl(M)
End Function

Function PushUnderLin(O$())
Push O, UnderLin(LasEle(O))
End Function

Function PushUnderLinDbl(O$())
Push O, UnderLinDbl(LasEle(O))
End Function

Function UnderLinzLines$(Lines$, Optional UnderLinChr$ = "-")
UnderLinzLines = Lines & vbCrLf & Dup("-", WdtzLines(Lines))
End Function

Function PushMsgUnderLin(O$(), M$)
Push O, M
Push O, UnderLin(M)
End Function
