Attribute VB_Name = "MDta_Dry_Col_Rmv"
Option Explicit
Function DryRmvC(A(), C) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI DryRmvC, AyeEleAt(Dr, C)
Next
End Function

Function DryRmvIxAy(A(), IxAy) As Variant()
Dim Dr
For Each Dr In Itr(A)
   Push DryRmvIxAy, AyeIxAy(Dr, IxAy)
Next
End Function

