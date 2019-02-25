Attribute VB_Name = "MVb_Lin_Term_Asg"
Option Explicit

Sub Asg2TRst(Lin, OT1, OT2, ORst)
AsgApAy Sy2TRst(Lin), OT1, OT2, ORst
End Sub

Sub Asg3TRst(Lin, OT1, OT2, OT3, ORst)
AsgApAy Sy3TRst(Lin), OT1, OT2, OT3, ORst
End Sub

Sub Asg4T(Lin, O1, O2, O3, O4)
AsgApAy Fst4Term(Lin), O1, O2, O3, O4
End Sub

Sub Asg4TRst(Lin, O1, O2, O3, O4, ORst)
AsgApAy Sy4TRst(Lin), O1, O2, O3, O4, ORst
End Sub

Sub AsgTRst(Lin, OT1, ORst)
AsgApAy SyTRst(Lin), OT1, ORst
End Sub

Sub AsgTT(Lin, O1, O2)
AsgApAy Sy2TRst(Lin), O1, O2
End Sub
