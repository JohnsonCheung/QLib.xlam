Attribute VB_Name = "QIde_B_IsLin"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Is_Hit."
Private Const Asm$ = "QIde"
Function IsLinPrp(Lin) As Boolean
IsLinPrp = MthKd(Lin) = "Property"
End Function

Private Sub Z_IsLinMth()
GoTo Z
Dim A$
A = "Function IsLinMth(A) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = IsLinMth(A)
    C
    Return
Z:
Dim L, O$()
For Each L In CSrc
    If IsLinMth(CStr(L)) Then
        PushI O, L
    End If
Next
Brw O
End Sub

Function HitCnstn(SrcLin, Cnstn$) As Boolean
HitCnstn = CnstnzL(SrcLin) = Cnstn
End Function

Function HitCnstnDic(SrcLin, Cnstn As Aset) As Boolean
HitCnstnDic = Cnstn.Has(CnstnzL(SrcLin))
End Function

Function HitShtMdy(ShtMdy$, ShtMthMdyAy$()) As Boolean
HitShtMdy = HitAy(IIf(ShtMdy = "", "Pub", ShtMdy), ShtMthMdyAy)
End Function

Function IsLinOptOrImplOrBlnk(Lin) As Boolean
IsLinOptOrImplOrBlnk = True
If IsLinOpt(Lin) Then Exit Function
If IsLinImpl(Lin) Then Exit Function
If Lin = "" Then Exit Function
IsLinOptOrImplOrBlnk = False
End Function

Function IsLinImpl(Lin) As Boolean
IsLinImpl = HasPfx(Lin, "Implements ")
End Function

Function IsLinOpt(Lin) As Boolean
If Not HasPfx(Lin, "Option ") Then Exit Function
Select Case True
Case _
    HasPfx(Lin, "Option Explicit"), _
    HasPfx(Lin, "Option Compare Text"), _
    HasPfx(Lin, "Option Compare Binary"), _
    HasPfx(Lin, "Option Compare Database")
    IsLinOpt = True
End Select

End Function

Function IsLinPubMth(Lin) As Boolean
Dim L$: L = Lin
Dim Mdy$: Mdy = ShfMdy(L): If Mdy <> "" And Mdy <> "Public" Then Exit Function
IsLinPubMth = TakMthKd(Lin) <> ""
End Function

Function IsLinMth(Lin) As Boolean
IsLinMth = MthKd(Lin) <> ""
End Function
Function IsLinMthNm(Lin, Nm) As Boolean
IsLinMthNm = Mthn(Lin) = Nm
End Function

Function IsLinEmn(A) As Boolean
IsLinEmn = HasPfx(RmvMdy(A), "Enum ")
End Function

Function IsLinTy(A) As Boolean
IsLinTy = HasPfx(RmvMdy(A), "Type ")
End Function

Function IsLinEmpSrc(A) As Boolean
IsLinEmpSrc = True
If HasPfx(A, "Option ") Then Exit Function
Dim L$: L = Trim(A)
If L = "" Then Exit Function
IsLinEmpSrc = False
End Function

Function IsLinSngTerm(Lin) As Boolean
IsLinSngTerm = InStr(Trim(Lin), " ") = 0
End Function

Function IsLinDD(Lin) As Boolean
IsLinDD = Fst2Chr(LTrim(Lin)) = "--"
End Function


