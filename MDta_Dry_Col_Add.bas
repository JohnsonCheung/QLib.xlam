Attribute VB_Name = "MDta_Dry_Col_Add"
Option Explicit
    
Function AddColzDry(Dry(), C) As Variant()
Dim O(): O = Dry
Dim ToU&
    ToU = NColzDry(Dry)
Dim J&, Dr
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    Dr(ToU) = C
    O(J) = Dr
    J = J + 1
Next
AddColzDry = O
End Function

Function AddColzDry3C(A(), C1, C2, C3) As Variant()
Dim U%, R&, Dr, O()
O = A
U = NColzDry(A) + 2
For Each Dr In Itr(A)
    ReDim Preserve Dr(U)
    Dr(U) = C3
    Dr(U - 1) = C2
    Dr(U - 2) = C1
    O(R) = Dr
    R = R + 1
Next
AddColzDry3C = O
End Function

Function AddColzDryBy(Dry(), Optional ByNCol% = 1) As Variant()
Dim NewU&
    NewU = NColzDry(Dry) + ByNCol - 1
Dim O()
    Dim UDry&: UDry = UB(Dry)
    O = AyReSzU(O, UDry)
    Dim J&
    For J = 0 To UDry
        O(J) = AyReSzU(Dry(J), NewU)
    Next
AddColzDryBy = O
End Function
Function AddColzDryC(Dry(), C) As Variant()
If Sz(Dry) = 0 Then Exit Function
Dim O(): O = AddColzDryBy(Dry)
    Dim UCol%: UCol = UB(Dry(0))
    Dim J&
    For J = 0 To UB(Dry)
       O(J)(UCol) = C
    Next
AddColzDryC = O
End Function

Function AddColzDryCC(A, V1, V2) As Variant()
Dim O(): O = A
Dim ToU&
    ToU = NColzDry(A) + 1
Dim J&, Dr, I1%, I2%
I2 = ToU
I1 = I2 - 1
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    Dr(I1) = V1
    Dr(I2) = V2
    O(J) = Dr
    J = J + 1
Next
AddColzDryCC = O
End Function

Function InsColzDry3V(A(), V1, V2, V3) As Variant()
InsColzDry3V = DryInsAv(A, Av(V1, V2, V3))
End Function

Function InsColzDryV(A(), V, Optional At& = 0) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI InsColzDryV, AyInsItm(Dr, V, At)
Next
End Function

Function InsColzDry4V(A(), V1, V2, V3, V4) As Variant()
InsColzDry4V = DryInsAv(A, Av(V1, V2, V3, V4))
Sy
End Function

Function InsColzDry2V(A(), V1, V2) As Variant()
InsColzDry2V = DryInsAv(A, Av(V1, V2))
End Function


