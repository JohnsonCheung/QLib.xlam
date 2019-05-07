Attribute VB_Name = "QVb_Lin_Term_Asg"
Option Explicit
Private Const CMod$ = "MVb_Lin_Term_Asg."
Private Const Asm$ = "QVb"

Sub AsgN2tRst(Lin$, OT1$, OT2$, ORst$)
AsgAp SyzN2tRst(Lin), OT1, OT2, ORst
End Sub

Sub AsgN3tRst(Lin$, OT1$, OT2$, OT3$, ORst$)
AsgAp SyzN3TRst(Lin), OT1, OT2, OT3, ORst
End Sub

Sub AsgN4t(Lin$, O1$, O2$, O3$, O4$)
AsgAp Fst4Term(Lin), O1, O2, O3, O4
End Sub

Sub AsgN4tRst(Lin$, O1$, O2$, O3$, O4$, ORst$)
AsgAp SyzN4tRst(Lin), O1, O2, O3, O4, ORst
End Sub

Sub AsgTRst(Lin$, OT1, ORst)
AsgAp SyzTRst(Lin), OT1, ORst
End Sub

Sub AsgN2t(Lin$, O1, O2)
AsgAp SyzN2tRst(Lin), O1, O2
End Sub

Sub AsgT1FldLikAy(OT1, OFldLikAy$(), Lin$)
Dim Rst$
AsgTRst Lin, OT1, Rst
OFldLikAy = SyzSsLin(Rst)
End Sub

