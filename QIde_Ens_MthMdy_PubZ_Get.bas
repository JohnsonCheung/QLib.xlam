Attribute VB_Name = "QIde_Ens_MthMdy_PubZ_Get"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ens_MthMdy_PubZ_Get."
Private Const Asm$ = "QIde"
Private Const Ns$ = ""
Sub BrwMdNyzWiPubZ()
Brw MdNyzWiPubZ, "MdNyzWiPubZ"
End Sub
Function MdNyzWiPubZ() As String()
MdNyzWiPubZ = MdNyzWiPubzP(CPj)
End Function
Function MdNyzWiPubzP(P As VBProject) As String()
Dim C As VBComponent
For Each C In P.VBComponents
    If IsWiPubZMd(C.CodeModule) Then PushI MdNyzWiPubzP, C.Name
Next
End Function

Function MthLinAyzPubZM(M As CodeModule) As String()
Dim MthLin
For Each MthLin In Itr(MthLinAyzS(Src(M)))
    'If IsMthLinzPubZ(MthLin) Then PushI MthLinAyzPubZM, MthLin
Next
End Function

Function MthLinAyzPubZ() As String()
MthLinAyzPubZ = MthLinAyzPubZP(CPj)
End Function
Function MthLinAyzPubZP(P As VBProject) As String()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy MthLinAyzPubZP, MthLinAyzPubZM(C.CodeModule)
Next
End Function
Function IsWiPubZMd(M As CodeModule) As Boolean
Dim MthLin
For Each MthLin In Itr(MthLinAyzS(Src(M)))
    'If IsMthLinzPubZ(MthLin) Then IsWiPubZMd = True: Exit Function
Next
End Function

Function MthLinAyzPub(Src$()) As String()
Dim L
For Each L In Itr(Src)
    'If IsMthLinzPub(L) Then
        PushI MthLinAyzPub, L
    'End If
Next
End Function


