Attribute VB_Name = "MIde_Ens_CSub"
Option Explicit
Const DoczAy01$ = "Tag: FunNmSfx DtaTy.  It is an array with either 0 or 1 element."
Enum eActLin
    eeInsLin
    eeDltLin
    eeRplLin
End Enum
Const CMod$ = "MIde_Ens_CSub."
Private Type CModInf
    Lnx As Lnx
    IsUsingCMod As Boolean
    MdNm As String
    InsLno As Long
End Type
Private Type CSubInf
    IsUsingCSub As Boolean
    InsLno As Long
    Lnx As Lnx
    MthNm As String
End Type

Sub EnsCSubMd()
EnsCSubzMd CurMd
End Sub

Private Function ActMdAy01(A As CodeModule) As ActMd()
Dim S$():                 S = Src(A)
Dim Rg() As MthRg:       Rg = MthRgAy(S)
Dim CInf() As CSubInf: CInf = CSubInfAy(S, Rg) 'Should same# of ele as Rg
Dim MInf As CModInf:   MInf = CModInf(A, CInf)
Dim MAct() As ActLin:  MAct = CModAct(MInf)
Dim CAct() As ActLin:  CAct = CSubAct(CInf)
Dim OAy() As ActLin:    OAy = OyAdd(MAct, CAct)
If Si(OAy) > 0 Then
    Dim ActMd As New ActMd
    PushObj ActMdAy01, ActMd.Init(A, OAy)
End If
End Function

Sub Z_ActMdAyzPj()
Dim Pj As VBProject, Act() As ActMd, Ept() As ActMd
GoSub ZZ
Exit Sub
ZZ:
    Set Pj = CurPj
    GoTo Tst
Tst:
    Act = ActMdAyzPj(Pj)
    Brw LyzActMdAy(Act): Stop
    Return
End Sub
Function LyzActMdAy(A() As ActMd) As String()
If Si(A) = 0 Then Exit Function
Dim O$()
PushI O, A(0).Hdr
Dim J%
For J = 0 To UB(A)
    PushIAy O, A(J).ToLy
Next
LyzActMdAy = FmtAyT3(O)
End Function
Sub EnsCSubzMd(A As CodeModule, Optional Silent As Boolean)
Dim Act() As ActMd: Act = ActMdAy01(A)
If Si(Act) = 0 Then Exit Sub
MdMdy A, Act(0).ActLinAy, Silent
SavPj PjzMd(A)
End Sub
Private Sub Z_ActMdAy01()
Dim Md1 As CodeModule, Act() As ActMd
GoSub T0
Exit Sub
T0:
    Set Md1 = CurMd
    GoTo Tst
Tst:
    Act = ActMdAy01(Md1)
    Brw Act(0).ToLy
    Return
End Sub

Private Function ActMdAyzPj(Pj As VBProject) As ActMd()
If Pj.Protection = vbext_pp_locked Then Thw CSub, "Pj is locked", "Pj", Pj.Name
Dim C As VBComponent
For Each C In Pj.VBComponents
    PushObjAy ActMdAyzPj, ActMdAy01(C.CodeModule) '<===
Next
End Function

Sub EnsCSubzPj(A As VBProject, Optional Silent As Boolean)
PjMdy A, ActMdAyzPj(A)
End Sub

Sub EnsCSubPj(Optional Silent As Boolean)
EnsCSubzPj CurPj, Silent
End Sub

Private Function CModLin$(MdNm$)
CModLin = FmtQQ("Const CMod$ = ""?.""", MdNm)
End Function

Private Function CSubLin$(MthNm$)
CSubLin = FmtQQ("Const CSub$ = CMod & ""?""", MthNm)
End Function

Private Function CSubLnx(Src$(), FmIx&, ToIx&) As Lnx
Dim Ix
Dim O As New Lnx
For Ix = FmIx + 1 To ToIx - 1
    If HasPfx(Src(Ix), "Const CSub$") Then
        O.Ix = Ix
        O.Lin = Src(Ix)
        Set CSubLnx = O
        Exit Function
    End If
Next
Set CSubLnx = O
End Function

Private Function CSubInf(Src$(), B As MthRg) As CSubInf
With CSubInf
    Set .Lnx = CSubLnx(Src, B.FmIx, B.ToIx)
    .InsLno = InsLnoCSub(Src, B)
    .IsUsingCSub = IsUsingCSub(Src, B)
    .MthNm = B.MthNm
End With
End Function

Sub EnsCSub()
EnsCSubzMd CurMd
End Sub


Private Function CModInf(A As CodeModule, B() As CSubInf) As CModInf
With CModInf
.MdNm = MdNm(A)
Set .Lnx = CModLnx(A)
.InsLno = InsLnoCMod(A)
.IsUsingCMod = IsUsingCMod(B)
End With
End Function

Private Function CSubInfAy(Src$(), A() As MthRg) As CSubInf()
Dim U&: U = UB(A)
If U < 0 Then Exit Function
Dim O() As CSubInf
ReDim O(U) As CSubInf
Dim J%
For J = 0 To U
    O(J) = CSubInf(Src, A(J))
Next
CSubInfAy = O
End Function

Private Function ShouldInsCModLin(A As CModInf) As Boolean
With A
    Select Case True
    Case .IsUsingCMod And .Lnx.Lin = "": ShouldInsCModLin = True
    Case .IsUsingCMod And .Lnx.Lin <> CModLin(A.MdNm): ShouldInsCModLin = True
    End Select
End With
End Function

Private Function ShouldDltCModLin(A As CModInf) As Boolean
With A
    Select Case True
    Case .IsUsingCMod And .Lnx.Lin <> CModLin(A.MdNm): ShouldDltCModLin = True
    Case Not .IsUsingCMod And .Lnx.Lin <> "": ShouldDltCModLin = True
    End Select
End With
End Function

Private Function ShouldInsCSubLin(A As CSubInf) As Boolean
With A
    Select Case True
    Case .IsUsingCSub And .Lnx.Lin = "": ShouldInsCSubLin = True
    Case .IsUsingCSub And .Lnx.Lin <> CSubLin(A.MthNm): ShouldInsCSubLin = True
    End Select
End With
End Function

Private Function ShouldDltCSubLin(A As CSubInf) As Boolean
With A
    Select Case True
    Case .IsUsingCSub And .Lnx.Lin <> CSubLin(A.MthNm): ShouldDltCSubLin = True
    Case Not .IsUsingCSub And .Lnx.Lin <> "": ShouldDltCSubLin = True
    End Select
End With

End Function

Private Function CModAct(A As CModInf) As ActLin()
Dim IsUsing As Boolean, OldLin$, NewLin$, Lno&
IsUsing = A.IsUsingCMod
Lno = A.Lnx.Lno
OldLin = A.Lnx.Lin
NewLin = CModLin(A.MdNm)
If ShouldIns(IsUsing, OldLin, NewLin) Then PushObj CModAct, ActLin(eeInsLin, Lno, NewLin)
If ShouldDlt(IsUsing, OldLin, NewLin) Then PushObj CModAct, ActLin(eeDltLin, Lno, OldLin)
End Function

Private Function ActLin(Act As eActLin, Lno&, Lin$) As ActLin
Set ActLin = New ActLin
ActLin.Init Act, Lin, Lno
End Function

Private Function CSubInfSz%(A() As CSubInf)
On Error GoTo X
CSubInfSz = UBound(A) + 1
X:
End Function

Private Function CSubInfUB%(A() As CSubInf)
CSubInfUB = CSubInfSz(A) - 1
End Function

Private Function CSubAct(A() As CSubInf) As ActLin()
Dim J%
For J = 0 To CSubInfUB(A)
    PushObjAy CSubAct, CSubActzSng(A(J))
Next
End Function

Private Function InsLnoCSub&(Src$(), A As MthRg)
Dim J&
For J = A.FmIx + 1 To A.ToIx - 1
    If LasChr(Src(J - 1)) <> "_" Then
        InsLnoCSub = J + 1
        Exit Function
    End If
Next
'No need to throw error, just exit it returns 0
'Thw CSub, "Cannot find Lno where to insert CSub of a given method", "MthNm MthLy", A.MthNm, AywFT(Src, A.FmIx, A.ToIx)
End Function

Private Function InsLnoCMod&(A As CodeModule)
InsLnoCMod = FstLnozAftOptMd(A)
End Function

Private Function CSubActzSng(A As CSubInf) As ActLin()
Dim IsUsing As Boolean, OldLin$, NewLin$, Lno&, InsLno&
IsUsing = A.IsUsingCSub
InsLno = A.InsLno
Lno = A.Lnx.Lno
OldLin = A.Lnx.Lin
NewLin = CSubLin(A.MthNm)
If ShouldIns(IsUsing, OldLin, NewLin) Then PushObj CSubActzSng, ActLin(eeInsLin, InsLno, NewLin)
If ShouldDlt(IsUsing, OldLin, NewLin) Then PushObj CSubActzSng, ActLin(eeDltLin, Lno, OldLin)
'StopEr ActMdEr(CSubActzSng)
End Function

Function CvActMd(A) As ActMd
Set CvActMd = A
End Function

Private Function CModLnx(Md As CodeModule) As Lnx
Dim J%, L$
For J = 1 To Md.CountOfDeclarationLines
    L = Md.Lines(J, 1)
    If HasPfx(L, "Const CMod$") Then
        Set CModLnx = Lnx(J - 1, L)
        Exit Function
    End If
Next
Dim O As New Lnx
Set CModLnx = O
End Function

Private Function IsUsingCSub(Src$(), A As MthRg) As Boolean
Const CSub$ = CMod & "IsUsingCSub"
Dim J%
With A
    For J = .FmIx + 1 To .ToIx - 1
        If HasSubStr(Src(J), "CSub, ") Then
            IsUsingCSub = True
            Exit Function
        End If
    Next
End With
End Function

Private Function IsUsingCMod(A() As CSubInf) As Boolean
Dim J%
For J = 0 To CSubInfUB(A)
    If A(J).IsUsingCSub Then
        IsUsingCMod = True
        Exit Function
    End If
Next
End Function
