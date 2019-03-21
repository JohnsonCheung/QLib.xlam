Attribute VB_Name = "MIde_Ens_MthMdy"
Option Explicit
Const CMod$ = "MIde_Ens_Mdy."
Function MthLinzEnsPrv$(MthLin)
Const CSub$ = CMod & "MthLinzEnsPrv"
If Not IsMthLin(MthLin) Then Thw CSub, "Given MthLin is not MthLin", "Lin", MthLin
MthLinzEnsPrv = "Private " & RmvMdy(MthLin)
End Function

Function MthLinzEnsPub$(MthLin)
Const CSub$ = CMod & "MthLinzEnsPub"
If Not IsMthLin(MthLin) Then Thw CSub, "Given MthLin is not MthLin", MthLin
MthLinzEnsPub = RmvMdy(MthLin)
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
MdMdy A, ActLinAyzEnsPrvZ(A)
End Sub
Function ActLinAyzEnsPrvZ(A As CodeModule) As ActLin()

End Function
Function LnoAyzPubZ(A As CodeModule) As Long()
Dim L, J&
For Each L In Itr(Src(A))
    J = J + 1
    If IsMthLinzPub(L) Then
        PushI LnoAyzPubZ, J
    End If
Next
End Function

Function LnoItrzPubZ(A As CodeModule)
Asg Itr(LnoAyzPubZ(A)), LnoItrzPubZ
End Function

Sub EnsPubzMd(A As CodeModule)
MdMdy A, ActLinAyzEnsPubZ(A)
End Sub

Function ActLinAyzEnsPubZ(A As CodeModule) As ActLin()

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

Function ActLinzEnsPub(A, MthNm) As ActLin

End Function

Sub EnsPub(A As CodeModule, MthNm$)
MdyLin A, ActLinzEnsPub(A, MthNm)
End Sub

Function ActLinzEnsPrv(A As CodeModule, MthNm) As ActLin

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

