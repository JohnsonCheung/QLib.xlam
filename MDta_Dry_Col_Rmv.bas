Attribute VB_Name = "MDta_Dry_Col_Rmv"
Option Explicit
Function RmvColzDryC(A(), C) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI RmvColzDryC, AyeEleAt(Dr, C)
Next
End Function

Function RmvColzDryIxAy(Dry(), IxAy) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
   Push RmvColzDryIxAy, AyeIxAy(Dr, IxAy)
Next
End Function

