Attribute VB_Name = "MVb_Lin_Vy"
Option Explicit

Function ShfVy(OLin$, Lblss$) As Variant() ' _
'Return Ay, which is _
'   Same sz as Lblss-cnt _
'   Ay-ele will be either string of boolean _
'   Each element is corresponding to terms-lblss _
'Update OLin _
'   if the term match, it will removed from OLin _
'Lblss: is in *LLL ?LLL or LLL _
'   *LLL is always first at beginning, mean the value OLin has not lbl _
'   ?LLL means the value is in OLin is using LLL => it is true, _
'   LLL  means the value in OLin is LLL=VVV _
'OLin is _
'   VVV VVV=LLL [VVV=L L]
Dim L, Ay$(), O()
Ay = TermAy(OLin)
For Each L In Itr(SySsl(Lblss))
    Select Case FstChr(L)
    Case "*":
        Select Case SndChr(L)
        Case "?":  If Si(Ay) = 0 Then Thw CSub, "Must BoolLbl of Lblss not found in OLin", "Must-Bool-Lbl OLin Lblss", L, OLin, Lblss
                   PushI O, CBool(ShfFstEle(Ay))
        Case Else: PushI O, ShfFstEle(Ay)
        End Select
    Case "?": PushI O, ShfBool(Ay, L)
    Case Else: PushI O, ShfTxt(Ay, L)
    End Select
Next
ShfVy = O
End Function

Private Function ShfTxtOpt(OAy$(), Lbl) As StrOpt
If Si(OAy) = 0 Then Exit Function
Dim S$: S = ShfTxt(OAy, Lbl)
If S = "" Then ShfTxtOpt = SomStr(S)
End Function

Private Function ShfBool(OAy$(), Lbl)
If Si(OAy) = 0 Then Exit Function
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

Private Function ShfTxt(OAy$(), Lbl)
If Si(OAy) = 0 Then Exit Function
'Return either string or ""
Dim I, J%, Ay$()
Ay = OAy
For Each I In Itr(Ay)
    With Brk2(I, "=")
        If .s1 = Lbl Then
            ShfTxt = .s2
            OAy = AyeEleAt(Ay, J)
            Exit Function
        End If
    End With
    J = J + 1
Next
End Function

Private Sub Z_ShfVy()
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
    Act = ShfVy(Lin, Lblss)
    C Act, Ept
    Return
End Sub

