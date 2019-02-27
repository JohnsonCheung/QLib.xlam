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
        Case "?":  With ShfBoolOpt(Ay, L)
                        If Not .Som Then Thw CSub, FmtQQ("Must BoolLbl(?) not found in Lin(?). Lblss=(?)", L, Lin, Lblss)
                        PushI O, .Bool
                   End With
        Case Else: With ShfTxtOpt(Ay, L)
                        If Not .Som Then Thw CSub, FmtQQ("Must TxtLbl(?) not found in Lin(?), Lblss=(?)", L, Lin, Lblss)
                        PushI O, .Str
                   End With
        End Select
    Case "?": PushI VyzLinLbl, ShfBool(Ay, L)
    Case Else: PushI VyzLinLbl, ShfTxt(Ay, L)
    End Select
Next
End Function

Private Function ShfBoolOpt(OAy$(), Lbl) As BoolRslt
If Sz(OAy) = 0 Then Exit Function

End Function

Private Function ShfTxtOpt(OAy$(), Lbl) As StrRslt
If Sz(OAy) = 0 Then Exit Function

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
GoSub Cas1
GoSub Cas2
GoSub Cas3
Exit Sub
Dim Lin$, Lblss$, EptLin$
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
    Act = VyzLinLbl(Lin, Lblss)
    C
    Ass Lin = EptLin
    Return
End Sub

