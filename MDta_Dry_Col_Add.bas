Attribute VB_Name = "MDta_Dry_Col_Add"
Option Explicit
    
Function DryAddColz(Dry(), C) As Variant()
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
DryAddColz = O
End Function

Function DryAddColz3C(A(), C1, C2, C3) As Variant()
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
DryAddColz3C = O
End Function

Function DryAddColzBy(Dry(), Optional ByNCol% = 1) As Variant()
Dim NewU&
    NewU = NColzDry(Dry) + ByNCol - 1
Dim O()
    Dim UDry&: UDry = UB(Dry)
    O = AyReSzU(O, UDry)
    Dim J&
    For J = 0 To UDry
        O(J) = AyReSzU(Dry(J), NewU)
    Next
DryAddColzBy = O
End Function

Function DryAddColzC(Dry(), C) As Variant()
If Si(Dry) = 0 Then Exit Function
Dim O(): O = DryAddColzBy(Dry)
    Dim UCol%: UCol = UB(Dry(0))
    Dim J&
    For J = 0 To UB(Dry)
       O(J)(UCol) = C
    Next
DryAddColzC = O
End Function

Function DryAddColzCC(Dry(), V1, V2) As Variant()
DryAddColzCC = DryAddColzAv(Dry, Av(V1, V2))
End Function

Function DryAddColzAv(Dry(), Av()) As Variant()
Dim O(): O = Dry
Dim ToU&
    ToU = NColzDry(Dry) + 1
Dim J&, Dr, I1%, I2%
I2 = ToU
I1 = I2 - 1
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    PushAy Dr, Av
    O(J) = Dr
    J = J + 1
Next
DryAddColzAv = O
End Function

Function DryInsColzAv(Dry(), Av()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DryInsColzAv, AyAdd(Dr, Av)
Next
End Function
Function DryInsColz3V(Dry(), V1, V2, V3) As Variant()
DryInsColz3V = DryInsColzAv(Dry, Av(V1, V2, V3))
End Function

Function DryInsColzV(A(), V, Optional At& = 0) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI DryInsColzV, AyInsItm(Dr, V, At)
Next
End Function

Function DryInsColz4V(A(), V1, V2, V3, V4) As Variant()
DryInsColz4V = DryInsColzAv(A, Av(V1, V2, V3, V4))
End Function

Function DryInsColz2V(A(), V1, V2) As Variant()
DryInsColz2V = DryInsColzAv(A, Av(V1, V2))
End Function


