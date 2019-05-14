Attribute VB_Name = "QIde_Ens__Mdy"
Option Explicit
Private Const CMod$ = "MIde_Ens__Mdy."
Private Const Asm$ = "QIde"

Sub MdyLins(A As CodeModule, B As Mdygs)
Dim J&
For J = 0 To B.N - 1
    MdyLin A, B.Ay(J)
Next
End Sub

Sub MdyLin(A As CodeModule, B As Mdyg)
Select Case B.Act
Case EmMdyg.EiIns: InsLinzMI A, B.Ins
Case EmMdyg.EiDlt: DltLinzMD A, B.Dlt
Case EmMdyg.EiRpl: RplLinzMR A, B.Rpl
Case EmMdyg.EiNop
Case Else
Stop
    Thw CSub, "Invalid Mdyg.Act" ', "Vdt-ActLin.Act ActLin", "Mdyg.Act", LB.ToStr
End Select
End Sub

Function MdySrcLin(Src$(), B As ActLin) As String()
Dim O$()
Select Case B.Act
Case EmLinAct.EiInsLin
    MdySrcLin = CvSy(AyInsEle(Src, B.Lin, B.Ix))
Case EmLinAct.EiDltLin
    If Src(B.Ix) <> B.Lin Then Stop
    MdySrcLin = AyeEleAt(Src, B.Ix)
Case Else
    Thw CSub, "Invalid ActLin.Act, it should EiInsLin or EiDltLin, ", "ActLin", B.ToStr
End Select
End Function

Sub MdyPj(P As VBProject, B As Mdygs)
Dim I
'For Each I In Itr(B)
'    With CvActMd(I)
'        Debug.Print Mdn(.Md)
'        MdyMd .Md, .ActLiny, Silent
'    End With
'Next
End Sub

Private Sub Z_MdySrc()
Dim Mdy() As ActLin
GoSub ZZ2
Exit Sub
ZZ2:
    Dim M, J%
    For Each M In MdItr(CPj)
        If J > 10 Then Exit For
        Dim A() As ActMd
        'A = EnsgCModSub(CvMd(M))
        If Si(A) > 0 Then
            Brw MdySrc(Src(CvMd(M)), A(0).ActLiny), Mdn(CvMd(M))
        End If
    Next
    Return
Tst:
    Return

End Sub
Private Sub Z_FmtEnsCSubzMd()
Dim Md As CodeModule
'GoSub ZZ1
GoSub ZZ2
Exit Sub
ZZ1:
    Set Md = CMd
    GoTo Tst
ZZ2:
    Dim M
    For Each M In MdItr(CPj)
        Dim O$()
        O = FmtEnsCSubzMd(CvMd(M))
        If Si(O) > 0 Then Brw O, Mdn(CvMd(M))
    Next
    Return
Tst:
    Act = FmtEnsCSubzMd(Md)
    Brw Act
    Return
End Sub
Function FmtEnsCSubzMd(A As CodeModule) As String()
Dim ActMdAy01() As ActMd
'ActMdAy01 = EnsgCModSub(A)
Select Case Si(ActMdAy01)
Case 0:
Case 1: FmtEnsCSubzMd = FmtMdySrc(Src(A), ActMdAy01(0).ActLiny)
Case Else: Thw CSub, "Err in ActMdAy01zEnsCS: Should return Ay of si 1 or 0"
End Select
End Function
Function FmtMdygs(B As Mdygs, Src$()) As String()
Dim N%: N = NDig(Si(Src))
Dim J&, Middle$, Lno$, Lin, O$()
Dim DltDic As New Dictionary
    For J = 0 To UB(B)
        If B(J).Act = EiDltLin Then
            DltDic.Add B(J).Ix, B(J)
        End If
    Next
Dim InsDic As New Dictionary
    For J = 0 To UB(B)
        If B(J).Act = EiInsLin Then
            InsDic.Add B(J).Ix, B(J)
        End If
    Next

For J = 0 To UB(Src)
    Dim IsIns As Boolean, IsDlt As Boolean
    Dim InsLin, DltLin
        If DltDic.Exists(J) Then
            IsDlt = True
            'DltLin = CvActLin(DltDic(J)).Lin
        Else
            IsDlt = False
            DltLin = ""
        End If
        
        If InsDic.Exists(J) Then
            IsIns = True
            'InsLin = CvActLin(InsDic(J)).Lin
        Else
            IsIns = False
            InsLin = ""
        End If
    
    Lno = AlignR(J + 1, N)
    Lin = Src(J)
    If IsDlt Then If DltLin <> Lin Then Stop
    Middle = IIf(IsDlt, " <<<<< ", "       ")
    If IsIns Then
        PushI O, Lno & " >>>>> " & InsLin
    End If
    PushI O, Lno & Middle & Lin
Next
FmtMdySrc = O
End Function
Function SrcAftMdy(Src$(), B As Mdygs) As String()
ThwIf_Er ErOf_MdgyLins(B), CSub, "Error in Mdygs", "Mdygs Src", LyzActLiny(B), Src
Dim O$(), J&
O = Src
For J = UB(B) To 0 Step -1
    O = SrcAftMdy(O, B(J))
Next
SrcAftMdy = O
End Function

Sub MdyMd(A As MdygMd)
Dim J&
'For J = 0 To A.N - 1
'    MdyLin A, A.Ay(J)
'Next
'Dim NewLines$: NewLines = JnCrLf(MdySrc(Src(A), B)): 'Brw NewLines: Stop
'RplMd A, NewLines
End Sub

Private Sub PushActEr(O$(), Msg$, Ix, Cur As ActLin, Las As ActLin)
Dim Nav(3)
Nav(0) = "CurIx Las Cur"
Nav(1) = Ix
Nav(2) = Las.ToStr
Nav(3) = Cur.ToStr
PushIAy O, LyzMsgNav(Msg, Nav) '<-----------
End Sub

Private Function ErzActLinCurLas(Ix, Cur As ActLin, Las As ActLin) As String()
Dim O$(), A() As ActLin
With Cur
    Select Case True
    Case .Lno = 0
        PushActEr O, "Lno cannot be zero", Ix, Cur, Las
    Case Not HasPfx(.Lin, "Const C")
        PushActEr O, "ActMd.Lin must with pfx Const C", Ix, Cur, Las
    Case Else
        Select Case Las.Lno
        Case Is > .Lno: PushActEr O, "ActMd not in order", Ix, Cur, Las
        Case .Lno:
            Select Case True
            Case .Act = Las.Act
                PushActEr O, "Two same Lno should not same have 'IsIns'", Ix, Cur, Las
            Case .Act = EiInsLin
                PushActEr O, "For same line, the Later one (CurLno) should be delete, but now it is insert", Ix, Cur, Las
            End Select
        Case Else
        End Select
    End Select
End With
ErzActLinCurLas = O
End Function

Private Function ErzMdygs(A As Mdygs) As String()
Dim Ix%, Las As ActLin, Msg$, Cur Asg ActLin
If Si(A) <= 1 Then Exit Function
Set Las = A(0)
For Ix = 1 To UB(A)
    Set Cur = A(Ix)
    PushIAy ErzActLiny, ErzActLinCurLas(Ix, Cur, Las) '<===
    Set Las = Cur
Next
End Function

Private Function LyzActLiny(A() As ActLin) As String()
Dim J%
For J = 0 To UB(A)
    PushI LyzActLiny, A(J).ToStr
Next
End Function
