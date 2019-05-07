Attribute VB_Name = "QDta_Col_RmvCol"
Option Explicit
Private Const CMod$ = "MDta_Dry_Col_Rmv."
Private Const Asm$ = "QDta"
Function RmvColzDry(Dry(), C&) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI RmvColzDry, AyeEleAt(Drv, C)
Next
End Function

Function RmvColzDryIxAy(Dry(), IxAy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
   Push RmvColzDryIxAy, AyeIxAy(Dr, IxAy)
Next
End Function

