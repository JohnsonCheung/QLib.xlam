Attribute VB_Name = "QIde_Ens_MthMdy"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_MthMdy."
Function MthLinzEnsprv$(MthLin)
Const CSub$ = CMod & "MthLinzEnsprv"
If Not IsMthLin(MthLin) Then Thw CSub, "Given MthLin is not MthLin", "Lin", MthLin
MthLinzEnsprv = "Private " & RmvMdy(MthLin)
End Function

Function MthLinzEnspub$(MthLin)
Const CSub$ = CMod & "MthLinzEnspub"
If Not IsMthLin(MthLin) Then Thw CSub, "Given MthLin is not MthLin", MthLin
MthLinzEnspub = RmvMdy(MthLin)
End Function

Sub EnsMdPrvZ()
EnsPrvZzMd CMd
End Sub

Sub EnsPjPrvZ()
EnsPrvZzP CPj
End Sub

Sub EnsPrvZzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsPrvZzMd C.CodeModule
Next
End Sub

Sub EnsPubMd()
EnsPubzMd CMd
End Sub

Sub EnsPrvZzMd(A As CodeModule)
'MdyMd A, MdygsOfEnsPrvZ(A)
End Sub

Function LnoAyOfPubZ(A As CodeModule) As Long()
Dim L, J&
For Each L In Itr(Src(A))
    J = J + 1
    'If IsMthLinzPub(L) Then
        PushI LnoAyOfPubZ, J
    'End If
Next
End Function

Function LnoItrOfPubZ(A As CodeModule)
Asg Itr(LnoAyOfPubZ(A)), LnoItrOfPubZ
End Function

Sub EnsPubzMd(A As CodeModule)
'MdyMd A, MdygsOfEnsPubZ(A)
End Sub


Function LnoItrPrvZ(A As CodeModule)

End Function

Sub EnsMdy(A As CodeModule, Mthn, Optional Mdy$)
End Sub

Sub EnsPrv(A As CodeModule, Mthn)
EnsMdy A, Mthn, "Private"
End Sub

Function EnsgPub(A As CodeModule, Mthn) As Mdyg
End Function

Sub EnsPub(A As CodeModule, Mthn)
'MdyLin EnsgPub(A, Mthn)
End Sub

Function EnsgPrv(A As CodeModule, Mthn) As Mdyg

End Function

Private Function MthLinzEnsMdy$(OldMthLin, ShtMdy$)
Const CSub$ = CMod & "MthLinzEnsMdy"
Dim L$: L = RmvMdy(OldMthLin)
    Select Case ShtMdy
    Case "Pub", "": MthLinzEnsMdy = L
    Case "Prv":     MthLinzEnsMdy = "Private " & L
    Case "Frd":     MthLinzEnsMdy = "Friend " & L
    Case Else
        Thw CSub, "Given parameter [ShtMdy] must be ['' Pub Prv Frd]", "ShtMdy", ShtMdy
    End Select
End Function


Private Sub Z_EnsMdy()
Dim M As CodeModule
Dim Mthn
Dim Mdy$
'--
Set M = CMd
Mthn = "Z_A"
Mdy = "Prv"
GoSub Tst
Exit Sub
Tst:
    EnsMdy M, Mthn, Mdy
    Return
End Sub

Private Sub ZZ()
MIde_EnsMdy:
End Sub

