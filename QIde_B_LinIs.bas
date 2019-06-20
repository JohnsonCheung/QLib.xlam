Attribute VB_Name = "QIde_B_LinIs"
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
With Mthn3zL(Lin)
    Select Case True
    Case .Nm = "":
    Case .ShtMdy = "Pub": IsLinPubMth = True
    End Select
End With
End Function

