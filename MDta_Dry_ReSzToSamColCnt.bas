Attribute VB_Name = "MDta_Dry_ReSzToSamColCnt"
Option Explicit
Function DryReSzToSamColCnt(Dry()) As Variant()
If Si(Dry) = 0 Then Exit Function
Dim U&: U = UB(Dry(0))
Dim NeedReSz As Boolean
    Dim IU&, Dr
    For Each Dr In Dry
        IU = UB(Dr)
        If U <> IU Then
            NeedReSz = True
            U = Max(U, IU)
        End If
    Next
Dim O(): O = Dry
If NeedReSz Then
    Dim J&
    For Each Dr In O
        If U <> IU Then
            ReDim Preserve Dr(U)
            O(J) = Dr
        End If
        J = J + 1
    Next
End If
DryReSzToSamColCnt = O
End Function

