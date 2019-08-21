Attribute VB_Name = "QVb_F_TyDfn"
Option Explicit
Option Compare Text
Public Const FFoTyDfn$ = "Mdn Nm Ty Mem Rmk"
Private Sub Z_DoTyDfnP()
BrwDrs DoTyDfnP
End Sub

Function TyDfnNyP() As String()
TyDfnNyP = TyDfnNyzP(CPj)
End Function

Function TyDfnNyzP(P As VBProject) As String()
Dim L: For Each L In RmkLy(SrczP(P))
    PushNB TyDfnNyzP, TyDfnNm(L)
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
Dim Dy(): Dy = DyoTyDfn(RmkLy(S))
Dim ODy(): ODy = InsColzDy(Dy, C.Name)
DoTyDfnzCmp = Drs(FoTyDfn, ODy)
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
SqoTyDfnzP = SqzDy(DyoTyDfn(RmkLy(SrczP(P))))
End Function

Private Function DyoTyDfn(RmkLy$()) As Variant()
':DyoTyDfn: :Dyo-Nm-Ty-Mem-Rmk #Dyo-TyDfn# ! Fst-Lin must be :nn: :dd #mm# !rr
'                                          ! Rst-Lin is !rr
'                                          ! must term: nn dd mm, all of them has no spc
'                                          ! opt      : mm rr
'                                          ! :xx:     : should uniq in pj
Dim Gp(): Gp = GpoTyDfn(RmkLy)
Dim IGp: For Each IGp In Itr(Gp)
    PushSomSi DyoTyDfn, DroTyDfn(IGp)
Next
End Function

Private Function GpoTyDfn(RmkLy$()) As Variant()
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
GpoTyDfn = O
End Function

Private Function IsLinTyDfnRmk(Lin) As Boolean
If FstChr(Lin) <> "'" Then Exit Function
If FstChr(LTrim(RmvFstChr(Lin))) <> "!" Then Exit Function
IsLinTyDfnRmk = True
End Function

Private Sub Z_DroTyDfn()
Dim RmkLy
GoSub ZZ
Exit Sub
ZZ:
    RmkLy = Sy("':Cell: :SCell-or-:WCell")
    Dmp DroTyDfn(RmkLy)
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

Private Function IsLinTyDfn(Lin) As Boolean
Dim Nm$, Dfn$, T3$, Rst$
Asg3TRst Lin, Nm, Dfn, T3, Rst
IsLinTyDfn = IsTyDfn(Nm, Dfn, T3, Rst)
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
Private Function DroTyDfn(RmkLy) As Variant()
'Assume: Fst Lin is ':nn: :dd [#mm#] [!rr]
'        Rst Lin is '                 !rr
If Not IsArray(RmkLy) Then Stop
If Si(RmkLy) = 0 Then Exit Function
Dim Nm$, Dfn$, Mem$, Rmk$
AsgTyDfn RmkLy(0), Nm, Dfn, Mem, Rmk
'Rmk = Add RmkzVbRmkLy(RmvFstEle(RmkLy))
DroTyDfn = Array(Nm, Dfn, Mem, Rmk)
End Function
