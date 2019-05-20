Attribute VB_Name = "QIde_Ens__Mdy"
Option Compare Text
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
Case EmMdyg.EiIns: InsLinzM A, B.Ins
Case EmMdyg.EiDlt: DltLinzM A, B.Dlt
Case EmMdyg.EiNop
Stop
    Thw CSub, "Invalid Mdyg.Act" ', "Vdt-Mdyg.Act Mdyg", "Mdyg.Act", LB.ToStr
End Select
End Sub

Private Function InsLinzS(Src$(), A As Insg) As String()
InsLinzS = AyInsEle(Src, A.Lin, A.Lno - 1)
End Function

Private Function DltLinzS(Src$(), A As Dltg) As String()
If Src(A.Lno - 1) <> A.Lin Then Stop
DltLinzS = AyeEleAt(Src, A.Lno - 1)
End Function

Function MdySrc(Src$(), M As Mdygs) As String()
Dim J&, O$()
O = Src
For J = 0 To M.N - 1
    O = MdySrczSM(O, M.Ay(J))
Next
End Function

Function MdySrczSM(Src$(), M As Mdyg) As String()
Select Case M.Act
Case EmMdyg.EiIns: MdySrczSM = InsLinzS(Src, M.Ins)
Case EmMdyg.EiDlt: MdySrczSM = DltLinzS(Src, M.Dlt)
Case Else:         Thw CSub, "Invalid Mdyg.Act, it should EiInsLin or EiDltLin, ", "Mdyg.Act", M.Act
End Select
End Function

Sub MdyPj(P As VBProject, B As Mdygs)
Dim I
'For Each I In Itr(B)
'    With CvActMd(I)
'        Debug.Print Mdn(.Md)
'        MdyMd .Md, .Mdygs, Silent
'    End With
'Next
End Sub

Private Sub Z_MdySrc()
Dim Mdy() As Mdyg
GoSub ZZ2
Exit Sub
ZZ2:
    Dim M, J%
    For Each M In MdItr(CPj)
        If J > 10 Then Exit For
        Dim A As Mdygs
        'A = EnsgCModSub(CvMd(M))
        'If Si(A) > 0 Then
        '   Brw MdySrc(Src(CvMd(M)), A(0).Mdygs), Mdn(CvMd(M))
        'End If
    Next
    Return
Tst:
    Return

End Sub

Function FmtMdMdyg(Src$(), M As Mdygs) As String()
Dim N%: N = NDig(Si(Src))
Dim J&, Middle$, Lno$, Lin, O$()
Dim DltDic As New Dictionary
    For J = 0 To M.N - 1
        If M.Ay(J).Act = EiDlt Then
            'DltDic.Add B(J).Ix, B(J)
        End If
    Next
Dim InsDic As New Dictionary
    For J = 0 To M.N - 1
        If M.Ay(J).Act = EiIns Then
            'InsDic.Add B(J).Ix, B(J)
        End If
    Next

For J = 0 To UB(Src)
    Dim IsIns As Boolean, IsDlt As Boolean
    Dim InsLin, DltLin
        If DltDic.Exists(J) Then
            IsDlt = True
            'DltLin = CvMdyg(DltDic(J)).Lin
        Else
            IsDlt = False
            DltLin = ""
        End If
        
        If InsDic.Exists(J) Then
            IsIns = True
            'InsLin = CvMdyg(InsDic(J)).Lin
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
FmtMdMdyg = O
End Function

Function SrcAftMdyg(Src$(), B As Mdygs) As String()
ThwIf_Er ErzMdygs(B), CSub
Dim O$(), J&
O = Src
For J = B.N - 1 To 0 Step -1
    O = MdySrczSM(O, B.Ay(J))
Next
SrcAftMdyg = O
End Function

Sub MdyMd(A As RplgMd)
Dim J&
'For J = 0 To A.N - 1
'    MdyLin A, A.Ay(J)
'Next
'Dim NewLines$: NewLines = JnCrLf(MdySrc(Src(A), B)): 'Brw NewLines: Stop
'RplMd A, NewLines
End Sub

Private Sub PushMdygEr(O$(), Msg$, Ix, Cur As Mdyg, Las As Mdyg)
Dim Nav(3)
Nav(0) = "CurIx Las Cur"
Nav(1) = Ix
Nav(2) = FmtMdyg(Las)
Nav(3) = FmtMdyg(Cur)
PushIAy O, LyzMsgNav(Msg, Nav) '<-----------
End Sub
Private Function LnozMdyg&(A As Mdyg)
With A
    Select Case True
    Case .Act = EiDlt: LnozMdyg = .Ins.Lno
    Case .Act = EiDlt: LnozMdyg = .Dlt.Lno
    End Select
End With
End Function
Private Function ErzMdygCurLas(Ix, Cur As Mdyg, Las As Mdyg) As String()
Dim O$(), A() As Mdyg
With Cur
    Select Case True
    Case LnozMdyg(Cur) = 0
        'PushActEr O, "Lno cannot be zero", Ix, Cur, Las
    'Case Not HasPfx(.Lin, "Const C")
        'PushActEr O, "ActMd.Lin must with pfx Const C", Ix, Cur, Las
    Case Else
'        Select Case Las.Lno
'        Case Is > .Lno: PushActEr O, "ActMd not in order", Ix, Cur, Las
'        Case .Lno:
'            Select Case True
'            Case .Act = Las.Act
'                PushActEr O, "Two same Lno should not same have 'IsIns'", Ix, Cur, Las
'            Case .Act = EiInsLin
'                PushActEr O, "For same line, the Later one (CurLno) should be delete, but now it is insert", Ix, Cur, Las
'            End Select
'        Case Else
        End Select
'    End Select
End With
ErzMdygCurLas = O
End Function

Private Function ErzMdygs(A As Mdygs) As String()
', "Error in Mdygs", "Mdygs Src", FmtMdygs(B), Src

Dim Ix%, Las As Mdyg, Msg$, Cur As Mdyg
'If Si(A) <= 1 Then Exit Function
'Las = A(0)
'For Ix = 1 To UB(A)
    'Cur = A(Ix)
    PushIAy ErzMdygs, ErzMdygCurLas(Ix, Cur, Las) '<===
    Las = Cur
'Next
End Function

Private Function FmtMdygs(A As Mdygs) As String()
Dim J&
For J = 0 To A.N - 1
    PushI FmtMdygs, FmtMdyg(A.Ay(J))
Next
End Function
