Attribute VB_Name = "MxTyDfn"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxTyDfn."
Public Const FFoTyDfn$ = "Mdn Nm Ty Mem Rmk"
Private Sub Z_DoTyDfnP()
BrwDrs DoTyDfnP
End Sub

Function TyDfnNyP() As String()
TyDfnNyP = TyDfnNyzP(CPj)
End Function

Function TyDfnNyzP(P As VBProject) As String()
Dim L: For Each L In VbRmk(SrczP(P))
    PushNB TyDfnNyzP, TyDfnNm(L)
Next
End Function

Private Function IsLinOkTyDfn(L) As Boolean
Dim Nm$, Dfn$, T3$, Rst$
Asg3TRst L, Nm, Dfn, T3, Rst
IsLinOkTyDfn = IsTyDfn(Nm, Dfn, T3, Rst)
End Function

Private Function IsLinTyDfn(L) As Boolean
If FstChr(L) <> "'" Then Exit Function
Dim T$: T = T1(L)
If Fst2Chr(T) <> "':" Then Exit Function
If LasChr(T) <> ":" Then Exit Function
IsLinTyDfn = True
End Function

Function IsLinNkTyDfn(L) As Boolean
IsLinNkTyDfn = Not IsLinOkTyDfn(L)
End Function

Function NkTyDfnLy(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLinTyDfn(L) Then
        If IsLinNkTyDfn(L) Then
            PushI NkTyDfnLy, L
        End If
    End If
Next
End Function

Function TyDfnNm$(Lin)
Dim T$: T = T1(Lin)
If T = "" Then Exit Function
If Fst2Chr(T) <> "':" Then Exit Function
If LasChr(T) <> ":" Then Exit Function
TyDfnNm = RmvFstChr(T)
End Function

Function DoTyDfnP() As Drs
':DoTyDfn: :Drs-Nm-Ty-Mem-Rmk
DoTyDfnP = DoTyDfnzP(CPj)
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

Function WsoTyDfn() As Worksheet
Set WsoTyDfn = WsoTyDfnzP(CPj)
End Function

Function WsoTyDfnzP(P As VBProject) As Worksheet
Dim O As New Worksheet
Set O = NewWs("TyDfn")
'RgzSq DocSqzP(P), A1zWs(O)
Stop
FmtWsoTyDfn O
Set WsoTyDfnzP = O
End Function

Private Sub FmtWsoTyDfn(WsoTyDfn As Worksheet)

End Sub

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

Private Function IsLinTyDfnRmk(Lin) As Boolean
If FstChr(Lin) <> "'" Then Exit Function
If FstChr(LTrim(RmvFstChr(Lin))) <> "!" Then Exit Function
IsLinTyDfnRmk = True
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

Private Function IsTyDfn(Nm$, Dfn$, T3$, Rst$) As Boolean
Select Case True
Case Fst2Chr(Nm) <> "':"
Case LasChr(Nm) <> ":"
Case FstChr(Dfn) <> ":"
Case T3 <> "" And Not HasPfxSfx(T3, "#", "#") And FstChr(T3) <> "!"
Case Else: IsTyDfn = True
End Select
End Function


Private Sub AsgTyDfn(FstRmkLin, ONm$, ODfn$, OMem$, ORmk$)
Dim T3$, Rst$
Asg3TRst FstRmkLin, ONm, ODfn, T3, Rst
If Not IsTyDfn(ONm, ODfn, T3, Rst) Then
    Debug.Print ONm, ODfn, T3, Rst
    Stop
    Debug.Print "DroTyDfn: Fst lin of @RmkLy is not ':nn: :dd [#mm#] [!rr].  FstLin=[" & FstRmkLin & "]"
    Exit Sub
End If
ONm = RmvFstChr(ONm)
If HasPfxSfx(T3, "#", "#") Then OMem = T3
ORmk = Trim(RmvPfx(Rst, "!"))
End Sub

Private Function RmkzTyDfnRmkLy$(TyDfnRmkLy$())
Dim R$, O$()
Dim L: For Each L In Itr(TyDfnRmkLy)
    If FstChr(L) = "'" Then
        Dim A$: A = LTrim(RmvFstChr(L))
        If FstChr(A) = "!" Then
            PushNB O, LTrim(RmvFstChr(A))
        End If
    End If
Next
RmkzTyDfnRmkLy = JnCrLf(O)
End Function

Private Function DroTyDfn(Rmk$(), Mdn$) As Variant()
'Assume: Fst Lin is ':nn: :dd [#mm#] [!rr]
'        Rst Lin is '                 !rr
If Si(Rmk) = 0 Then Exit Function
Dim Nm$, Dfn$, Mem$, RmkLin$
AsgTyDfn Rmk(0), Nm, Dfn, Mem, RmkLin
Dim Rmkl1$: Rmkl1 = AddNB(RmkLin, Rmkl(CvSy(RmvFstEle(Rmk))))
DroTyDfn = Array(Mdn, Nm, Dfn, Mem, Rmkl1)
End Function
