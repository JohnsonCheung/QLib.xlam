Attribute VB_Name = "QVb_Lnoss_Lnoss"
Type AySubAy
    Ay As Variant
    SubAy As Variant
End Type
Function AySubAy(Ay, SubAy) As AySubAy
ThwIfNotAy Ay, CSub
ThwIfNotAy SubAy, CSub
AySubAy.Ay = Ay
AySubAy.SubAy = SubAy
End Function
Function LnossSy(A As AySubAy) As String()
Dim Itm
For Each Itm In Itr(A.SubAy)
    PushI LnossSy, Lnoss(A.Ay, Itm)
Next
End Function
Private Function Lnoss$(Ay, Itm)
Dim O&(), V, J&
For Each V In Ay
    J = J + 1
    If V = Itm Then PushI O, J
Next
Lnoss = JnSpc(O)
End Function

