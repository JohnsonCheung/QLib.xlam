Attribute VB_Name = "QVb_Dta_Rpt"
Option Explicit
Option Compare Text
Enum EmRpt
    EiRptOnly
    EiUpdAndRpt
    EiUpdOnly
    EiPushOnly
    EiUpdAndPush
End Enum

Enum EmHdr
    EiNoHdr
    EiWiHdr
End Enum

Function StrzRpt$(Rpt As EmRpt)
Dim O$
Select Case True
Case Rpt = EiRptOnly: O = "*RptOnly"
Case Rpt = EiUpdAndRpt: O = "*UpdAndRpt"
Case Rpt = EiUpdOnly: O = "*UpdOnly"
Case Else: O = "EmRptEr(" & Rpt & ")"
End Select
StrzRpt = O
End Function
Function IsPushzRpt(Rpt As EmRpt) As Boolean
Select Case True
Case Rpt = EiPushOnly, Rpt = EiUpdAndPush: IsPushzRpt = True
End Select
End Function
Function IsRptzRpt(Rpt As EmRpt) As Boolean
Select Case True
Case Rpt = EiUpdAndRpt, Rpt = EiRptOnly: IsRptzRpt = True
End Select
End Function
Function IsUpdzRpt(Rpt As EmRpt) As Boolean
Select Case True
Case Rpt = EiUpdAndRpt, Rpt = EiUpdOnly: IsUpdzRpt = True
End Select
End Function

