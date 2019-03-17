Attribute VB_Name = "MIde_Mth_Lin_Is_Hit"
Option Explicit
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
For Each L In SrcMd
    If IsMthLin(CStr(L)) Then
        PushI O, L
    End If
Next
Brw O
End Sub
Function HitConstNm(SrcLin, ConstNm) As Boolean
HitConstNm = ConstNmzSrcLin(SrcLin) = ConstNm
End Function

Function HitMthLin(MthLin, B As WhMth) As Boolean
HitMthLin = HitMthNm3(MthNm3(MthLin), B)
End Function

Function HitMthNm3(A As MthNm3, B As WhMth) As Boolean
If A.IsEmp Then Exit Function
If IsNothing(B) Then HitMthNm3 = True: Exit Function
If B.IsEmp Then HitMthNm3 = True: Exit Function
Select Case True
Case A.Nm = "":
Case IsNothing(B): HitMthNm3 = True
Case B.IsEmp:
Case Not HitNm(A.Nm, B.WhNm)
Case Not HitShtMdy(A.ShtMdy, B.ShtMdyAy)
Case Not HitAy(A.ShtKd, B.ShtKdAy)
Case Else: HitMthNm3 = True
End Select
End Function

Function HitShtMdy(ShtMdy$, ShtMdyAy$()) As Boolean
HitShtMdy = HitAy(IIf(ShtMdy = "", "Pub", ShtMdy), ShtMdyAy)
End Function

Function IsMthLinzPubZ(Lin) As Boolean
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
With MthNm3(Lin)
    Select Case True
    Case .Nm <> "":
    Case .MthMdy <> "" And .MthMdy <> "Public":
    Case Else: IsMthLinzPub = True
    End Select
End With
End Function

