Attribute VB_Name = "QIde_Mth_Lin_Is_Hit"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Is_Hit."
Private Const Asm$ = "QIde"
Function IsPrpLin(Lin) As Boolean
IsPrpLin = MthKd(Lin) = "Property"
End Function

Private Sub Z_IsLinzMth()
GoTo ZZ
Dim A$
A = "Function IsLinzMth(A) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = IsLinzMth(A)
    C
    Return
ZZ:
Dim L, O$()
For Each L In CSrc
    If IsLinzMth(CStr(L)) Then
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

Function IsOptLinOfOrImplzOrBlank(Lin) As Boolean
IsOptLinOfOrImplzOrBlank = True
If IsOptLin(Lin) Then Exit Function
If IsImplLin(Lin) Then Exit Function
If Lin = "" Then Exit Function
IsOptLinOfOrImplzOrBlank = False
End Function
Function IsImplLin(Lin) As Boolean
IsImplLin = HasPfx(Lin, "Implements ")
End Function

Function IsOptLin(Lin) As Boolean
If Not HasPfx(Lin, "Option ") Then Exit Function
Select Case True
Case _
    HasPfx(Lin, "Option Explicit"), _
    HasPfx(Lin, "Option Compare Text"), _
    HasPfx(Lin, "Option Compare Binary"), _
    HasPfx(Lin, "Option Compare Database")
    IsOptLin = True
End Select

End Function

Function IsPMthLin(Lin) As Boolean
With Mthn3zL(Lin)
    Select Case True
    Case .Nm = "":
    Case .ShtMdy = "Pub": IsPMthLin = True
    End Select
End With
End Function

