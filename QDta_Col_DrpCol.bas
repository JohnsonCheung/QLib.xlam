Attribute VB_Name = "QDta_Col_DrpCol"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Dry_Col_Rmv."
Private Const Asm$ = "QDta"
Function DrpColzDry(Dry(), C&) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI DrpColzDry, AyeEleAt(Drv, C)
Next
End Function

Function DrpColzDryIxy(Dry(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
   Push DrpColzDryIxy, AyeIxy(Dr, Ixy)
Next
End Function

