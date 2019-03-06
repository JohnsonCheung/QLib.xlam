Attribute VB_Name = "MVb_Lin_Term_Asg"
Option Explicit

Sub Asg2TRst(Lin, OT1, OT2, ORst)
AsgAp Syz2TRst(Lin), OT1, OT2, ORst
End Sub

Sub Asg3TRst(Lin, OT1, OT2, OT3, ORst)
AsgAp Syz3TRst(Lin), OT1, OT2, OT3, ORst
End Sub

Sub Asg4T(Lin, O1, O2, O3, O4)
AsgAp Fst4Term(Lin), O1, O2, O3, O4
End Sub

Sub Asg4TRst(Lin, O1, O2, O3, O4, ORst)
AsgAp Syz4TRst(Lin), O1, O2, O3, O4, ORst
End Sub

Sub AsgTRst(Lin, OT1, ORst)
AsgAp SyzTRst(Lin), OT1, ORst
End Sub

Sub AsgTT(Lin, O1, O2)
AsgAp Syz2TRst(Lin), O1, O2
End Sub

Sub AsgT1FldLikAy(OT1$, OFldLikAy$(), Lin)
Dim Rst$
AsgTRst Lin, OT1, Rst
OFldLikAy = SySsl(Rst)
End Sub

