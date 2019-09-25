Attribute VB_Name = "MxAlignDy"
Option Explicit
Option Compare Text
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxAlignDy."
Function AlignDrWyAsLin$(Dr, WdtAy%())
'Ret : a lin by joing [ | ] and quoting [| * |] after aligng @Dr with @WdtAy. @@
AlignDrWyAsLin = QteJnzAsTLin(AlignDrWy(Dr, WdtAy))
End Function

Function AlignDrWyAsLy(Ay, WdtAy%()) As String()
Dim S, J&: For Each S In Ay
    PushI AlignDrWyAsLy, Align(S, WdtAy(J))
    J = J + 1
Next
End Function

Function AlignSqzW(Sq(), W%()) As Variant()
Dim O(): O = Sq
Dim IC%: For IC = 1 To UBound(Sq, 2)
    Dim Wdt%: Wdt = W(IC - 1)
    Dim IR&: For IR = 1 To UBound(Sq, 1)
        O(IR, IC) = Align(O(IR, IC), Wdt)
    Next
Next
AlignSqzW = O
End Function

Function AlignSq(Sq()) As Variant()
If Si(Sq) = 0 Then Exit Function
Dim C&, O(), NR&, NC&
NR = UBound(Sq, 1)
NC = UBound(Sq, 2)
ReDim O(1 To NR, 1 To NC)
For C = 1 To UBound(O, 2)
    AlignColzSq O, Sq, C, WdtzSqc(Sq, C)
Next
AlignSq = O
End Function

