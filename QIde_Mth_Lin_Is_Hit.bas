Attribute VB_Name = "QIde_Mth_Lin_Is_Hit"
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Is_Hit."
Private Const Asm$ = "QIde"
Function IsPrpLin(Lin) As Boolean
IsPrpLin = MthKd(Lin) = "Property"
End Function

Private Sub Z_IsMthLin()
GoTo ZZ
Dim A$
A = "Function IsMthLin(A, Optional B As WhMth) As Boolean"
Ept = True
GoSub Tst
Exit Sub
Tst:
    Act = IsMthLin(A)
    C
    Return
ZZ:
Dim L, O$()
For Each L In CurSrc
    If IsMthLin(CStr(L)) Then
        PushI O, L
    End If
Next
Brw O
End Sub

Function HitCnstn(SrcLin, Cnstn$) As Boolean
HitCnstn = CnstnzSrcLin(SrcLin) = Cnstn
End Function

Function HitCnstnDic(SrcLin, Cnstn As Aset) As Boolean
HitCnstnDic = Cnstn.Has(CnstnzSrcLin(SrcLin))
End Function

Function HitMthLin(Lin, B As WhMth) As Boolean
If Not IsMthLin(Lin) Then Exit Function
If IsNothing(B) Then HitMthLin = True: Exit Function
If B.IsEmp Then HitMthLin = True: Exit Function
If Not HitMthn3(Mthn3(Lin), B) Then Exit Function
HitMthLin = True
End Function

Function HitShtMdy(ShtMdy$, ShtMthMdyAy$()) As Boolean
HitShtMdy = HitAy(IIf(ShtMdy = "", "Pub", ShtMdy), ShtMthMdyAy)
End Function

Function IsMthLinzPubZ(Lin) As Boolean
With Mthn3(Lin)
    If FstChr(.Nm) <> "Z" Then Exit Function
    If .MthMdy <> "" Then
        If .MthMdy <> "Public" Then
            Exit Function
        End If
    End If
End With
IsMthLinzPubZ = True
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

Function IsMthLinzPub(Lin) As Boolean
With Mthn3(Lin)
    Select Case True
    Case .Nm <> "":
    Case .MthMdy <> "" And .MthMdy <> "Public":
    Case Else: IsMthLinzPub = True
    End Select
End With
End Function

