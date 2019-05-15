Attribute VB_Name = "QIde_Cmp_Itr"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Cmp_Itr."

Function ClsAyzP(P As VBProject) As CodeModule()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsCmpzCls(C) Then
        PushObj ClsAyzP, C
    End If
Next
End Function

Function ClsNyzP(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsCmpzCls(C) Then
        PushI ClsNyzP, C.Name
    End If
Next
End Function

Private Sub Z_CmpAyzP()
Dim Act() As VBComponent
Dim C, T As vbext_ComponentType
For Each C In CmpAyzP(CPj)
    T = CvCmp(C).Type
    If T <> vbext_ct_StdModule And T <> vbext_ct_ClassModule Then Stop
Next
End Sub

Function CmpAyzP(P As VBProject) As VBComponent()
If IsProtectzvInf(P) Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    If IsCmpzMd(C) Then
        PushObj CmpAyzP, C
    End If
Next
End Function

Function IsNoClsNoModPj(P As VBProject) As Boolean
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Type = vbext_ComponentType.vbext_ct_ClassModule Then Exit Function
    If C.Type = vbext_ComponentType.vbext_ct_StdModule Then Exit Function
Next
IsNoClsNoModPj = True
End Function
Function ModItrzP(P As VBProject)
Asg Itr(ModAyzP(P)), ModItrzP
End Function

Function ModAyzP(P As VBProject) As CodeModule()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    PushObj ModAyzP, C.CodeModule
Next
End Function

Function MdNyP() As String()
MdNyP = MdNyzP(CPj)
End Function
Function MdNyWiPrpV() As String()
MdNyWiPrpV = MdNyWiPrpzV(CVbe)
End Function
Function MdNyWiPrpzV(A As Vbe) As String()
Dim Mdn, I
For Each I In MdNyzV(A)
    Mdn = I
    If IsCmpzMdnWiPrp(Mdn) Then
        PushI MdNyWiPrpzV, Mdn
    End If
Next
End Function

Function IsCmpzMdnWiPrp(Mdn) As Boolean
Dim M As CodeModule: Set M = Md(Mdn)
Dim J&
For J = 1 To M.CountOfLines
    If IsPrpLin(M.Lines(J, 1)) Then IsCmpzMdnWiPrp = True: Exit Function
Next
End Function

Function MdNyV() As String()
MdNyV = MdNyzV(CVbe)
End Function

Function MdNyzM(Mthn) As String()
MdNyzM = MdnsetzM(Mthn).Sy
End Function

Function MdAyzNN(Mdnn$) As CodeModule()

End Function

Function MdAyzNy(MdNy$()) As CodeModule()
Dim N, P As VBProject
For Each N In Itr(MdNy)
    PushI MdAyzNy, MdzPN(P, N)
Next
End Function

Function MdAyzPm(PMth) As CodeModule()
MdAyzPm = MdAyzNy(MdNyzPm(PMth))
End Function

Function MdNyzPPm(P As VBProject, PMthn) As String()

End Function

Function MdNyzPm(PMthn) As String()
MdNyzPm = MdnsetzPm(PMthn).Sy
End Function

Function MdnsetzPm(PMthn) As Aset
Set MdnsetzPm = PMthnzRlMdnV.ParChd(PMthn)
End Function

Function MdnsetzM(Mthn) As Aset
Set MdnsetzM = MthnzRlMdnP.ParChd(Mthn)
End Function

Property Get PMthnzRlMdnV() As Rel
Set PMthnzRlMdnV = PMthnzRlMdnzV(CVbe)
End Property

Function PMthnzRlMdnzV(A As Vbe) As Rel
Stop
End Function

Function MdNyzP(P As VBProject) As String()
Dim C
For Each C In CmpItr(P)
    If IsCmpzMd(CvCmp(C)) Then
        PushI MdNyzP, C.Name
    End If
Next
End Function

Function MdNyzV(A As Vbe) As String()
Dim P As VBProject
For Each P In A.VBProjects
    PushIAy MdNyzV, MdNyzP(P)
Next
End Function

Function ModNy() As String()
ModNy = ModNyzP(CPj)
End Function

Function ModNyzP(P As VBProject) As String()
Dim C As VBComponent, O$()
For Each C In P.VBComponents
    If IsCmpzMod(C) Then PushI O, C.Name
Next
'ModNyzP = AywNmStr(O)
End Function

Private Sub Z_ClsNyzP()
DmpAy ClsNyzP(CPj)
End Sub



Private Sub Z_MdAy()
Dim O() As CodeModule
O = MdAyzP(CPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print Mdn(Md)
Next
End Sub

Private Sub Z_MdzPjny()
'DmpAy MdzPjny(CPj)
End Sub

Function CmpAyP() As VBComponent()
CmpAyP = CmpAyzP(CPj)
End Function
Function MdAy() As CodeModule()
MdAy = MdAyzP(CPj)
End Function

Function CmpItr(P As VBProject)
Asg Itr(CmpAyzP(P)), CmpItr
End Function

Function MdItr(P As VBProject)
Asg Itr(MdAyzP(P)), MdItr
End Function

Function MdAyzP(P As VBProject) As CodeModule()
MdAyzP = MdAyzC(CmpAyzP(P))
End Function

