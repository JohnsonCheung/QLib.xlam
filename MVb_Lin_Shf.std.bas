Attribute VB_Name = "MVb_Lin_Shf"
Option Explicit

Function ShfVal(OLin$, Lblss) As Variant()
'Return Ay, which is
'   Same sz as Lblss-cnt
'   Ay-ele will be either string of boolean
'   Each element is corresponding to terms-lblss
'Return OLin
'   if the term match, it will removed from OLin
'Lblss: is in *LLL ?LLL or LLL
'   *LLL is always first at beginning, mean the value OLin has not lbl
'   ?LLL means the value is in OLin is using LLL => it is true,
'   LLL  means the value in OLin is LLL=VVV
'OLin is
'   VVV VVV=LLL [VVV=L L]
Dim L, Ay$(), V
Ay = TermAy(OLin)
For Each L In Itr(SySsl(Lblss))
    Select Case FstChr(L)
    Case "*": PushI ShfVal, AyShf(Ay)
    Case "?": PushI ShfVal, ShfValBool(Ay, L)
    Case Else: PushI ShfVal, ShfValLblItm(Ay, L)
    End Select
Next
OLin = JnSpc(AyQuoteSqIf(Ay))
End Function
Private Function ShfValBool(OAy$(), Lbl) As Boolean
If Sz(OAy) = 0 Then Exit Function
Dim J%, L$, Ay$()
Ay = OAy
L = RmvFstChr(Lbl)
For J = 0 To UB(Ay)
    If Ay(J) = L Then
        ShfValBool = True
        OAy = AyeEleAt(Ay, J)
        Exit Function
    End If
Next
End Function

Private Function ShfValLblItm$(OAy$(), Lbl)
If Sz(OAy) = 0 Then Exit Function
'Return either string or ""
Dim I, J%, Ay$()
Ay = OAy
For Each I In Itr(Ay)
    With Brk2(I, "=")
        If .S1 = Lbl Then
            ShfValLblItm = .S2
            OAy = AyeEleAt(Ay, J)
            Exit Function
        End If
    End With
    J = J + 1
Next
End Function

Function ShfBktStr$(OLin$)
Dim O$
O = TakBetBkt(OLin): If O = "" Then Exit Function
ShfBktStr = O
OLin = TakAftBkt(OLin)
End Function
Function RmvChr$(S, ChrLis$) ' Rmv fst chr if it is in ChrLis
If HasSubStr(ChrLis, FstChr(S)) Then
    RmvChr = RmvFstChr(S)
Else
    RmvChr = S
End If
End Function
Function TakChr$(S, ChrLis$) ' Ret fst chr if it is in ChrLis
If HasSubStr(ChrLis, FstChr(S)) Then TakChr = FstChr(S)
End Function
Function ShfChr$(OLin, ChrList$)
Dim C$: C = TakChr(OLin, ChrList)
If C = "" Then Exit Function
ShfChr = C
OLin = Mid(OLin, 2)
End Function

Function ShfPfx(OLin, Pfx) As Boolean
If HasPfx(OLin, Pfx) Then
    OLin = RmvPfx(OLin, Pfx)
    ShfPfx = True
End If
End Function

Function ShfPfxSpc(OLin, Pfx) As Boolean
If HitPfxSpc(OLin, Pfx) Then
    OLin = Mid(OLin, Len(Pfx) + 2)
    ShfPfxSpc = True
End If
End Function

Private Sub Z_ShfVal()
GoSub Cas1
GoSub Cas2
GoSub Cas3
Exit Sub
Dim Lin$, Lblss, EptLin$
Cas1:
    Lin = "1 Req"
    EptLin = ""
    Lblss = "*XX ?Req"
    Ept = Array("1", True)
    GoTo Tst
Cas2:
    Lin = "A B C=123 D=XYZ"
    Lblss = "?B"
    Ept = Array(True)
    EptLin = "A C=123 D=XYZ"
    GoTo Tst
Cas3:
    Lin = "Txt VTxt=XYZ [Dft=A 1] VRul=123 Req"
    Lblss = "*Ty ?Req ?AlwZLen Dft VTxt VRul"
    Ept = Array("Txt", True, False, "A 1", "XYZ", "123")
    EptLin = ""
    GoTo Tst
Tst:
    Act = ShfVal(Lin, Lblss)
    C
    Ass Lin = EptLin
    Return
End Sub

Private Sub Z_ShfBktStr()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = ShfBktStr(A)
    C
    Ass A = Ept1
    Return
End Sub

Private Property Get Z_ShfPfx()
Dim O$: O = "AA{|}BB "
Ass ShfPfx(O, "{|}") = "AA"
Ass O = "BB "
End Property



Private Sub Z()
Z_ShfBktStr
Z_ShfVal
MVb_Lin_Shf:
End Sub
