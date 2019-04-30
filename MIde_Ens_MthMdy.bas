Attribute VB_Name = "MIde_Ens_MthMdy"
Option Explicit
Const CMod$ = "MIde_EnsMdy."
Function MthLinzEnsprv$(MthLin$)
Const CSub$ = CMod & "MthLinzEnsprv"
If Not IsMthLin(MthLin) Then Thw CSub, "Given MthLin is not MthLin", "Lin", MthLin
MthLinzEnsprv = "Private " & RmvMdy(MthLin)
End Function

Function MthLinzEnspub$(MthLin$)
Const CSub$ = CMod & "MthLinzEnspub"
If Not IsMthLin(MthLin) Then Thw CSub, "Given MthLin is not MthLin", MthLin
MthLinzEnspub = RmvMdy(MthLin)
End Function

Sub EnsMdPrvZ()
EnsPrvZzMd CurMd
End Sub

Sub EnsPjPrvZ()
EnsPrvZzPj CurPj
End Sub

Sub EnsPrvZzPj(A As VBProject)
Dim C As VBComponent
For Each C In A.VBComponents
    EnsPrvZzMd C.CodeModule
Next
End Sub

Sub EnsPubMd()
EnsPubzMd CurMd
End Sub

Sub EnsPrvZzMd(A As CodeModule)
MdyMd A, ActLinAyOfEnsPrvZ(A)
End Sub
Function ActLinAyOfEnsPrvZ(A As CodeModule) As ActLin()

End Function

Function LnoAyOfPubZ(A As CodeModule) As Long()
Dim L, J&
For Each L In Itr(Src(A))
    J = J + 1
    If IsMthLinzPub(L) Then
        PushI LnoAyOfPubZ, J
    End If
Next
End Function

Function LnoItrOfPubZ(A As CodeModule)
Asg Itr(LnoAyOfPubZ(A)), LnoItrOfPubZ
End Function

Sub EnsPubzMd(A As CodeModule)
MdyMd A, ActLinAyOfEnsPubZ(A)
End Sub

Function ActLinAyOfEnsPubZ(A As CodeModule) As ActLin()

End Function

Function LnoItrPrvZ(A As CodeModule)

End Function

Function ActLinzEnsMdy(A As CodeModule, MthNm, Mdy) As ActLin
End Function

Sub EnsMdy(A As CodeModule, MthNm$, Optional Mdy$)
MdyLin A, ActLinzEnsMdy(A, MthNm, Mdy)
End Sub

Sub EnsPrv(A As CodeModule, MthNm$)
EnsMdy A, MthNm, "Private"
End Sub

Function ActLinzEnspub(A, MthNm) As ActLin

End Function

Sub EnsPub(A As CodeModule, MthNm$)
MdyLin A, ActLinzEnspub(A, MthNm)
End Sub

Function ActLinzEnsprv(A As CodeModule, MthNm) As ActLin

End Function

Private Function MthLinzEnsMdy$(OldMthLin$, ShtMdy$)
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
Dim MthNm$
Dim Mdy$
'--
Set M = CurMd
MthNm = "Z_A"
Mdy = "Prv"
GoSub Tst
Exit Sub
Tst:
    EnsMdy M, MthNm, Mdy
    Return
End Sub

Private Sub Z()
MIde_EnsMdy:
End Sub

