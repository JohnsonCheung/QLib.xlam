Attribute VB_Name = "MDta_Dry_Col_Add"
Option Explicit
    
Function DryAddC(A, C) As Variant()
Dim O(): O = A
Dim ToU&
    ToU = NColDry(A)
Dim J&, Dr
For Each Dr In Itr(O)
    ReDim Preserve Dr(ToU)
    Dr(ToU) = C
    O(J) = Dr
    J = J + 1
Next
DryAddC = O
End Function

Function DryAddCol3C(A(), C1, C2, C3) As Variant()
Dim U%, R&, Dr, O()
O = A
U = NColDry(A) + 2
For Each Dr In Itr(A)
    ReDim Preserve Dr(U)
    Dr(U) = C3
    Dr(U - 1) = C2
    Dr(U - 2) = C1
    O(R) = Dr
    R = R + 1
Next
DryAddCol3C = O
End Function
Function DryIncCol(Dry(), Optional IncByN% = 1) As Variant()
Dim NewU&
    NewU = NColDry(Dry) + IncByN - 1
Dim O()
    Dim UDry&: UDry = UB(Dry)
    O = AyReSzU(O, UDry)
    Dim J&
    For J = 0 To UDry
        O(J) = AyReSzU(Dry(J), NewU)
    Next
DryIncCol = O
End Function
Function DryAddCol(Dry(), C) As Variant()
If Sz(Dry) = 0 Then Exit Function
Dim O(): O = DryIncCol(Dry)
    Dim UCol%: UCol = UB(Dry(0))
    Dim J&
    For J = 0 To UB(Dry)
       O(J)(UCol) = C
    Next
DryAddCol = O
End Function

Function DryAddCC(A, V1, V2) As Variant()
Dim O(): O = A
Dim ToU&
    ToU = NColDry(A) + 1
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
DryAddCC = O
End Function

Function DryIns3V(A(), V1, V2, V3) As Variant()
DryIns3V = DryInsAv(A, Av(V1, V2, V3))
End Function

Function DryInsV(A(), V, Optional At& = 0) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI DryInsV, AyInsItm(Dr, V, At)
Next
End Function

Function DryIns4V(A(), V1, V2, V3, V4) As Variant()
DryIns4V = DryInsAv(A, Av(V1, V2, V3, V4))
Sy
End Function

Function DryInsCC(A(), V1, V2) As Variant()
DryInsCC = DryInsAv(A, Av(V1, V2))
End Function


