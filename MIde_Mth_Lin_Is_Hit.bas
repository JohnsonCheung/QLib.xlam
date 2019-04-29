Attribute VB_Name = "MIde_Mth_Lin_Is_Hit"
Option Explicit
Function IsPrpLin(Lin$) As Boolean
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

Function HitConstNm(SrcLin$, ConstNm$) As Boolean
HitConstNm = ConstNmzSrcLin(SrcLin) = ConstNm
End Function

Function HitConstNmDic(SrcLin$, ConstNm As Aset) As Boolean
HitConstNmDic = ConstNm.Has(ConstNmzSrcLin(SrcLin))
End Function

Function HitMthLin(Lin$, B As WhMth) As Boolean
If Not IsMthLin(Lin) Then Exit Function
If IsNothing(B) Then HitMthLin = True: Exit Function
If B.IsEmp Then HitMthLin = True: Exit Function
If Not HitMthNm3(MthNm3(Lin), B) Then Exit Function
HitMthLin = True
End Function

Function HitMthNm3(A As MthNm3, B As WhMth) As Boolean
Select Case True
Case A.Nm = "":
Case IsNothing(B), B.IsEmp
    HitMthNm3 = True
Case _
    Not HitNm(A.Nm, B.WhNm), _
    Not HitShtMdy(A.ShtMdy, B.ShtMthMdyAy), _
    Not HitAy(A.ShtTy, B.ShtTyAy)
Case Else
    HitMthNm3 = True
End Select
End Function

Function HitShtMdy(ShtMdy$, ShtMthMdyAy$()) As Boolean
HitShtMdy = HitAy(IIf(ShtMdy = "", "Pub", ShtMdy), ShtMthMdyAy)
End Function

Function IsMthLinzPubZ(Lin$) As Boolean
With MthNm3(Lin)
    If FstChr(.Nm) <> "Z" Then Exit Function
    If .MthMdy <> "" Then
        If .MthMdy <> "Public" Then
            Exit Function
        End If
    End If
End With
IsMthLinzPubZ = True
End Function
Function IsOptLin(Lin$) As Boolean
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

Function IsMthLinzPub(Lin$) As Boolean
With MthNm3(Lin)
    Select Case True
    Case .Nm <> "":
    Case .MthMdy <> "" And .MthMdy <> "Public":
    Case Else: IsMthLinzPub = True
    End Select
End With
End Function

