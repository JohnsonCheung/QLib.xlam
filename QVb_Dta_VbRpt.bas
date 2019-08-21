Attribute VB_Name = "QVb_Dta_VbRpt"
Option Explicit
Option Compare Text
Enum EmUpd
    EiRptOnly
    EiUpdAndRpt
    EiUpdOnly
    EiPushOnly  ' Pushing to XX$()
    EiUpdAndPush
End Enum

Enum EmHdr
    EiNoHdr
    EiWiHdr
End Enum

Function EmUpdStr$(Upd As EmUpd)
Dim O$
Select Case True
Case Upd = EiRptOnly: O = "*RptOnly"
Case Upd = EiUpdAndRpt: O = "*UpdAndRpt"
Case Upd = EiUpdOnly: O = "*UpdOnly"
Case Else: O = "EmUpdEr(" & Upd & ")"
End Select
EmUpdStr = O
End Function
Function IsEmUpdRpt(Upd As EmUpd) As Boolean
Select Case True
Case Upd = EiPushOnly, Upd = EiUpdAndPush: IsEmUpdRpt = True
End Select
End Function

Function IsEmUpdUpd(Upd As EmUpd) As Boolean
Select Case True
Case Upd = EiUpdAndRpt, Upd = EiUpdOnly: IsEmUpdUpd = True
End Select
End Function

