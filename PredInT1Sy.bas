VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PredInT1Sy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IPred
Private Const CMod$ = "PredzInT1Sy."
Private A_T1Ay$()
Friend Sub Init(T1Ay$())
A_T1Ay = T1Ay
End Sub

Private Function IPred_Pred(V As Variant) As Boolean
Dim I, Lin, T1$
Lin = V
For Each I In A_T1Ay
    T1 = I
    If HasT1(Lin, T1) Then IPred_Pred = True: Exit Function
Next
End Function
