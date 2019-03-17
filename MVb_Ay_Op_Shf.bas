Attribute VB_Name = "MVb_Ay_Op_Shf"
Option Explicit

Private Sub Z_AyShf()
Dim Ay(), Exp, Act, ExpAyAft()
Ay = Array(1, 2, 3, 4)
Exp = 1
ExpAyAft = Array(2, 3, 4)
GoSub Tst
Exit Sub
Tst:
    Act = AyShf(Ay)
    Debug.Assert IsEq(Exp, Act)
    Debug.Assert IsEqAy(Ay, ExpAyAft)
    Return
End Sub

Private Sub Z_AyShfItm()
Dim OAy$(), Itm, EptOAy
Ept = "1"
Itm = "AA"
OAy = Sy("AA=1")
EptOAy = Sy("")
GoSub Tst
Exit Sub
Tst:
    Act = AyShfItm(OAy, Itm)
    C
    Ass IsEqAy(OAy, EptOAy)
    Return
End Sub

Private Sub Z_AyShfItmNy()
Dim A$(), ItmNy0
A = SySsl("Req Dft=ABC VTxt=kkk")
ItmNy0 = "Req ABC VTxt"
Ept = Array("Req", "", "kkk", Sy("Dft=ABC"))
GoSub Tst
Exit Sub
Tst:
    Act = AyShfItmNy(A, ItmNy0)
    C
    Return
End Sub

Function AyShf(OAy)
AyShf = OAy(0)
OAy = AyeFstEle(OAy)
End Function

Function AyShfFstNEle(OAy, N)
AyShfFstNEle = AywFstNEle(OAy, N)
OAy = AyeFstNEle(OAy, N)
End Function

Function AyShfItm$(OAy, Itm)
Dim J%
If FstChr(Itm) = "?" Then AyShfItm = AyShfQItm(OAy, RmvFstChr(Itm)): Exit Function
For J = 0 To UB(OAy)
    If HasPfx(OAy(J), Itm) Then
        AyShfItm = Brk(OAy(J), "=")(1)
        OAy = AyeEleAt(OAy, J)
        Exit Function
    End If
Next
End Function

Function AyShfItmEq(A, Itm$) As Variant()
Dim B$
    Dim Lik$
    Lik = Itm & "=*"
    B = FstEleLik(A, Lik)
If B = "" Then
    AyShfItmEq = Array("", A)
Else
    AyShfItmEq = Array(Trim(RmvPfx(B, Itm & "=")), AyeEleLik(A, Lik))
End If
End Function

Function AyShfItmNy(A$(), ItmNy0) As Variant()
Dim Ny$(), A1$()
    Ny = NyzNN(ItmNy0)
    A1 = A
Dim O() As Variant, Ay(), J%
ReDim O(Si(Ny))
For J = 0 To UB(Ny)
    Ay = AyShfItmEq(A1, Ny(J))
    O(J) = Ay(0)
    A1 = Ay(1)
Next
O(Si(Ny)) = Ay(1)
AyShfItmNy = O
End Function

Function AyShfQItm$(OAy, QItm)
Dim I, J%
For Each I In Itr(OAy)
    If QItm = I Then AyShfQItm = QItm: OAy = AyeEleAt(OAy, J): Exit Function
    J = J + 1
Next
End Function

Function AyShfStar(OAy, OItmy$()) As String()
Dim NStar%: NStar = AyNPfxStar(OItmy)
AyShfStar = AyShfFstNEle(OAy, NStar)
OItmy = AyeFstNEle(OItmy, NStar)
End Function


Private Sub Z()
Z_AyShf
Z_AyShfItm
Z_AyShfItmNy
MVb_AyShf:
End Sub
