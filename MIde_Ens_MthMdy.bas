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

Sub EnsPubZMd()
EnsPubZzMd CurMd
End Sub

Sub EnsPrvZzMd(A As CodeModule)
MdyMd ActMdzEnsPrvZ(A)
End Sub
Function ActMdzEnsPrvZ(A As CodeModule) As ActMd

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

Sub EnsPubZzMd(A As CodeModule)
MdyMd ActMdzEnsPubZ(A)
End Sub
Function ActMdzEnsPubZ(A As CodeModule) As ActMd

End Function
Function LnoItrPrvZ(A As CodeModule)

End Function
Function ActLinzEnsMthMdy(A As CodeModule, MthNm, Mdy) As ActLin
End Function
Sub EnsMdyzMth(A As CodeModule, MthNm$, Optional Mdy$)
MdyLin A, ActLinzEnsMthMdy(A, MthNm, Mdy)
End Sub
Sub EnsPrvZzMth(A As CodeModule, MthNm$)
EnsMdyzMth A, MthNm, "Private"
End Sub
Function ActLinzEnsPubMth(A, MthNm) As ActLin

End Function

Sub EnsPubzMth(A As CodeModule, MthNm$)
MdyLin A, ActLinzEnsPubMth(A, MthNm)
End Sub
Function ActLinzEnsPrvMth(A As CodeModule, MthNm) As ActLin

End Function
Sub EnsPrvzMth(A As CodeModule, MthNm$)
MdyLin A, ActLinzEnsPrvMth(A, MthNm)
End Sub

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


Private Sub Z_EnsMdyzMth()
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
    EnsMdyzMth M, MthNm, Mdy
    Return
End Sub

Private Sub Z()
MIde_EnsMdy:
End Sub

