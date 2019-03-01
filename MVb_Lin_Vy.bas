Attribute VB_Name = "MVb_Lin_Vy"
Option Explicit

Function VyzLinLbl(Lin$, Lblss$) As Variant()
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
Dim L, Ay$(), O()
Ay = TermAy(Lin)
For Each L In Itr(SySsl(Lblss))
    Select Case FstChr(L)
    Case "*":
        Select Case SndChr(L)
        Case "?":  If Sz(Ay) = 0 Then Thw CSub, "Must BoolLbl of Lblss not found in Lin", "Must-Bool-Lbl Lin Lblss", L, Lin, Lblss
                   PushI O, CBool(ShfFstEle(Ay))
        Case Else: PushI O, ShfFstEle(Ay)
        End Select
    Case "?": PushI O, ShfBool(Ay, L)
    Case Else: PushI O, ShfTxt(Ay, L)
    End Select
Next
VyzLinLbl = O
End Function

Private Function ShfTxtOpt(OAy$(), Lbl) As StrRslt
If Sz(OAy) = 0 Then Exit Function
Dim S$: S = ShfTxt(OAy, Lbl)
If S = "" Then ShfTxtOpt = StrRslt(S)
End Function

Private Function ShfBool(OAy$(), Lbl) As Boolean
If Sz(OAy) = 0 Then Exit Function
Dim J%, L$, Ay$()
Ay = OAy
L = RmvFstChr(Lbl)
For J = 0 To UB(Ay)
    If Ay(J) = L Then
        ShfBool = True
        OAy = AyeEleAt(Ay, J)
        Exit Function
    End If
Next
End Function

Private Function ShfTxt$(OAy$(), Lbl)
If Sz(OAy) = 0 Then Exit Function
'Return either string or ""
Dim I, J%, Ay$()
Ay = OAy
For Each I In Itr(Ay)
    With Brk2(I, "=")
        If .S1 = Lbl Then
            ShfTxt = .S2
            OAy = AyeEleAt(Ay, J)
            Exit Function
        End If
    End With
    J = J + 1
Next
End Function

Private Sub Z_VyzLinLbl()
GoSub T0
'GoSub T1
'GoSub T2
'GoSub T3
Exit Sub
Dim Lin$, Lblss$, Act(), Ept()
T0:
    Lin = "Loc Txt Req Dft=ABC AlwZLen [VTxt=Loc cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
    Lblss = EleLblss ' "*Fld *Ty ?Req ?AlwZLen Dft VTxt VRul TxtSz Expr"
    Ept = Array("Loc", "Txt", True, True, "ABC", "Loc cannot be blank", "IsNull([Loc]) or Trim(Loc)=''", "", "")
    GoTo Tst
T1:
    Lin = "1 Req"
    Lblss = "*XX ?Req"
    Ept = Array("1", True)
    GoTo Tst
T2:
    Lin = "A B C=123 D=XYZ"
    Lblss = "?B"
    Ept = Array(True)
    GoTo Tst
T3:
    Lin = "Txt VTxt=XYZ [Dft=A 1] VRul=123 Req"
    Lblss = "*Fld *Ty ?Req ?AlwZLen Dft VTxt VRul"
    Ept = Array("Txt", True, False, "A 1", "XYZ", "123")
    GoTo Tst
Tst:
    Act = VyzLinLbl(Lin, Lblss)
    C Act, Ept
    Return
End Sub

