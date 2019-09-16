Attribute VB_Name = "MxDoTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDoTyDfn."

Function DoTyDfn() As Drs
':DoTyDfn: :Drs-Mdn-Nm-Ty-Mem-Rmk
DoTyDfn = DoTyDfnzP(CPj)
End Function

Private Function DoTyDfnzP(P As VBProject) As Drs
Dim O As Drs
Dim C As VBComponent: For Each C In P.VBComponents
    O = AddDrs(O, DoTyDfnzCmp(C))
Next
DoTyDfnzP = O
End Function

Private Function DoTyDfnzCmp(C As VBComponent) As Drs
Dim S$(): S = Src(C.CodeModule)
Dim Dy(): Dy = DyoTyDfn(VbRmk(S), C.Name)
DoTyDfnzCmp = Drs(FoTyDfn, Dy)
End Function

Private Function FoTyDfn() As String()
FoTyDfn = SyzSS(FFoTyDfn)
End Function

Private Function SqoTyDfnzP(P As VBProject) As Variant()
SqoTyDfnzP = SqzDy(DyoTyDfnzP(P))
End Function

Private Function DyoTyDfnzP(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy DyoTyDfnzP, DyoTyDfnzM(C.CodeModule)
Next
End Function

Private Function DyoTyDfnzM(M As CodeModule) As Variant()
DyoTyDfnzM = DyoTyDfn(VbRmk(Src(M)), Mdn(M))
End Function

Private Function DyoTyDfn(VbRmk$(), Mdn$) As Variant()
':DyoTyDfn: :Dyo-Nm-Ty-Mem-VbRmk #Dyo-TyDfn# ! Fst-Lin must be :nn: :dd #mm# !rr
'                                          ! Rst-Lin is !rr
'                                          ! must term: nn dd mm, all of them has no spc
'                                          ! opt      : mm rr
'                                          ! :xx:     : should uniq in pj
Dim Gp(): Gp = LygoTyDfn(VbRmk)
Dim IGp: For Each IGp In Itr(Gp)
    PushSomSi DyoTyDfn, DroTyDfn(CvSy(IGp), Mdn)
Next
End Function

Private Function LygoTyDfn(RmkLy$()) As Variant()
Dim O()
Dim L: For Each L In Itr(RmkLy)
    Dim NFstLin%
    Dim Gp()
    Select Case True
    Case IsLinTyDfn(L)
        NFstLin = NFstLin + 1
        PushSomSi O, Gp
        Erase Gp
        PushI Gp, L
    Case IsLinTyDfnRmk(L)
        If Si(Gp) > 0 Then
            PushI Gp, L ' Only with Fst-Lin, the Rst-Lin will be use, otherwise ign it.
        End If
    Case Else
        PushSomSi O, Gp
        Erase Gp
    End Select
Next
LygoTyDfn = O
End Function

Private Sub Z_DroTyDfn()
Dim VbRmk$()
GoSub ZZ
Exit Sub
ZZ:
    VbRmk = Sy("':Cell: :SCell-or-:WCell")
    Dmp DroTyDfn(VbRmk, "Md")
    Return
End Sub

Private Function DroTyDfn(Rmk$(), Mdn$) As Variant()
'Assume: Fst Lin is ':nn: :dd [#mm#] [!rr]
'        Rst Lin is '                 !rr
If Si(Rmk) = 0 Then Exit Function
Dim NM$, Dfn$, Mem$, RmkLin$
Dim Dr(): Dr = DroTyDfnzL(Rmk(0))
Dim Rmkl1$: Rmkl1 = AddNB(RmkLin, Rmkl(CvSy(RmvFstEle(Rmk))))
DroTyDfn = Array(Mdn, NM, Dfn, Mem, Rmkl1)
End Function

Function DroTyDfnzL(FstTyDfnLin$) As Variant()
Dim T3$, Rst$
Dim ONm$, ODfn$, OMem$, ORmk$
Asg3TRst FstTyDfnLin, ONm, ODfn, T3, Rst
If Not IsTyDfn(ONm, ODfn, T3, Rst) Then
    Debug.Print ONm, ODfn, T3, Rst
    Stop
    Debug.Print "DroTyDfnzL: @FstTyDfnLin of @RmkLy is not ':nn: :dd [#mm#] [!rr].  FstLin=[" & FstTyDfnLin & "]"
    Exit Function
End If
ONm = RmvFstChr(ONm)
If HasPfxSfx(T3, "#", "#") Then OMem = T3
ORmk = Trim(RmvPfx(Rst, "!"))
DroTyDfnzL = Array(ONm$, ODfn$, OMem$, ORmk$)
End Function