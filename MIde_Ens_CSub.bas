Attribute VB_Name = "MIde_Ens_CSub"
Option Explicit
Enum eActLin
    eInsLin
    eDltLin
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

Private Function ActMd(A As CodeModule) As ActMd
Dim S$():                 S = Src(A)
Dim Rg() As MthRg:       Rg = MthRgAy(S)
Dim CInf() As CSubInf: CInf = CSubInfAy(S, Rg) 'Same# of ele as mMthRgAy
Dim MInf As CModInf:   MInf = CModInf(A, CInf)
Dim MAct  As ActLin:   MAct = CModAct(MInf)
Dim CAct() As ActLin:  CAct = CSubActAy(CInf)
Dim OAy() As ActLin:    OAy = IntozOy(OAy, ObjAddAy(MAct, CAct))
Set ActMd = New ActMd
ActMd.Init A, OAy
End Function

Sub EnsCSubzMd(A As CodeModule, Optional Silent As Boolean)
Dim Act As ActMd: Set Act = ActMd(A)
MdyMd Act, Silent
If Si(Act.ActLinAy) > 0 Then
    SavPj PjzMd(A)
End If
End Sub

Private Function ActPj(Pj As VBProject) As ActPj
If Pj.Protection = vbext_pp_locked Then Thw CSub, "Pj is locked", "Pj", Pj.Name
Dim C As VBComponent
Dim OAct() As ActMd
For Each C In Pj.VBComponents
    PushObj OAct, ActMd(C.CodeModule)
Next
Set ActPj = New ActPj
ActPj.Init OAct
End Function


Sub EnsCSubzPj(A As VBProject, Optional Silent As Boolean)
MdyPj ActPj(A), Silent
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

Private Function CSubLnx(Src$(), B As FTIx) As Lnx
Dim Ix
Dim O As New Lnx
For Ix = B.FmIx + 1 To B.ToIx - 1
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
    Set .Lnx = CSubLnx(Src, B)
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
.Lnx = CModLnx(A)
.InsLno = InsLnoCMod(A)
.IsUsingCMod = IsUsingCMod(B)
End With
End Function

Private Function CSubInfAy(Src$(), A() As MthRg) As CSubInf()
Dim N&: N = UB(A)
If N = 0 Then Exit Function
Dim O() As CSubInf
ReDim O(N - 1) As CSubInf
Dim J%
For J = 0 To N - 1
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

Private Function CModAct(A As CModInf) As ActLin
Dim IsUsing As Boolean, OldLin$, NewLin$, Lno&
IsUsing = A.IsUsingCMod
Lno = A.Lnx.Lno
OldLin = A.Lnx.Lin
NewLin = CModLin(A.MdNm)
If ShouldIns(IsUsing, OldLin, NewLin) Then PushObj CModAct, ActLin(eInsLin, Lno, NewLin)
If ShouldDlt(IsUsing, OldLin, NewLin) Then PushObj CModAct, ActLin(eDltLin, Lno, OldLin)
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
CSubInfUB = CSubInfSz(A)
End Function

Private Function CSubActAy(A() As CSubInf) As ActLin()
Dim J%, O() As ActMd
If CSubInfSz(A) = 0 Then Exit Function
ReDim O(CSubInfUB(A))
For J = 0 To CSubInfUB(A)
    O(J) = CSubActSng(A(J))
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
End Function

Private Function InsLnoCMod&(A As CodeModule)
InsLnoCMod = FstLnozAftOptMd(A)
End Function

Private Function CSubActSng(A As CSubInf) As ActLin()
Dim IsUsing As Boolean, OldLin$, NewLin$, Lno&, InsLno&
IsUsing = A.IsUsingCSub
InsLno = A.InsLno
Lno = A.Lnx.Lno
OldLin = A.Lnx.Lin
NewLin = CSubLin(A.MthNm)
If ShouldIns(IsUsing, OldLin, NewLin) Then PushObj CSubActSng, ActLin(eInsLin, Lno, NewLin)
If ShouldDlt(IsUsing, OldLin, NewLin) Then PushObj CSubActSng, ActLin(eDltLin, Lno, OldLin)
'StopEr ActMdEr(CSubActSng)
End Function

Private Function CvActMd(A) As ActMd
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
