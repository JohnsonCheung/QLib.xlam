Attribute VB_Name = "MIde_Ens__Mdy"
Option Explicit
Sub MdyLinAy(A As CodeModule, B() As ActLin)
ThwErMsg ErzActLinAy(B), CSub, "Error in ActLinAy", "ActLinAy Md", LyzActLinAy(B), MdNm(A)
Dim O$(), J&
For J = UB(B) To 0 Step -1
    MdyLin A, B(J)
Next
End Sub

Sub MdyLin(A As CodeModule, B As ActLin)
Select Case B.Act
Case eActLin.eeInsLin
    A.InsertLines B.Lno, B.Lin  '<== Inserted
    Inf CSub, "Lin is inserted", "Lno Lin", B.Lno, B.Lin
Case eActLin.eeDltLin
    Dim ActualLin$
    ActualLin = A.Lines(B.Lno, 1)
    If A.Lines(B.Lno, 1) <> B.Lin Then Thw CSub, "To MdLin to be deleted is not expected", "Lin-Expected-To-Delete Lin-Actual-In-Md Lno", B.Lin, ActualLin, B.Lno
    A.DeleteLines B.Lno, 1  '<== Deleted
    Inf CSub, "MdLin is deleted", "Lno Lin", B.Lno, ActualLin
Case Else
    Thw CSub, "Invalid ActLin.Act", "Vdt-ActLin.Act ActLin", "eeDltIn eeInsLin", B.ToStr
End Select
End Sub

Function SrcMdyLin(Src$(), B As ActLin) As String()
Dim O$()
Select Case B.Act
Case eActLin.eeInsLin
    SrcMdyLin = CvSy(AyInsEle(Src, B.Lin, B.Ix))
Case eActLin.eeDltLin
    If Src(B.Ix) <> B.Lin Then Stop
    SrcMdyLin = AyeEleAt(Src, B.Ix)
Case Else
    Thw CSub, "Invalid ActLin.Act, it should eeInsLin or eeDltLin, ", "ActLin", B.ToStr
End Select
End Function

Function PjMdy(A As VBProject, B() As ActMd, Optional Silent As Boolean) As VBProject
Dim I
For Each I In Itr(B)
    With CvActMd(I)
        Debug.Print MdNm(.Md)
        MdMdy .Md, .ActLinAy, Silent
    End With
Next
Set PjMdy = A
End Function

Private Sub Z_SrcMdy()
Dim Mdy() As ActLin
GoSub ZZ2
Exit Sub
ZZ2:
    Dim M, J%
    For Each M In MdItr(CurPj)
        If J > 10 Then Exit For
        Dim A() As ActMd
        A = ActMdAy01zEnsCSub(CvMd(M))
        If Si(A) > 0 Then
            Brw SrcMdy(Src(CvMd(M)), A(0).ActLinAy), MdNm(CvMd(M))
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
    Set Md = CurMd
    GoTo Tst
ZZ2:
    Dim M
    For Each M In MdItr(CurPj)
        Dim O$()
        O = FmtEnsCSubzMd(CvMd(M))
        If Si(O) > 0 Then Brw O, MdNm(CvMd(M))
    Next
    Return
Tst:
    Act = FmtEnsCSubzMd(Md)
    Brw Act
    Return
End Sub
Function FmtEnsCSubzMd(A As CodeModule) As String()
Dim ActMdAy01() As ActMd
ActMdAy01 = ActMdAy01zEnsCSub(A)
Select Case Si(ActMdAy01)
Case 0:
Case 1: FmtEnsCSubzMd = FmtSrcMdy(Src(A), ActMdAy01(0).ActLinAy)
Case Else: Thw CSub, "Err in ActMdAy01zEnsCS: Should return Ay of si 1 or 0"
End Select
End Function
Function FmtSrcMdy(Src$(), B() As ActLin) As String()
Dim N As Byte: N = NDig(Si(Src))
Dim J&, Middle$, Lno$, Lin$, O$()
Dim DltDic As New Dictionary
    For J = 0 To UB(B)
        If B(J).Act = eeDltLin Then
            DltDic.Add B(J).Ix, B(J)
        End If
    Next
Dim InsDic As New Dictionary
    For J = 0 To UB(B)
        If B(J).Act = eeInsLin Then
            InsDic.Add B(J).Ix, B(J)
        End If
    Next

For J = 0 To UB(Src)
    Dim IsIns As Boolean, IsDlt As Boolean
    Dim InsLin$, DltLin$
        If DltDic.Exists(J) Then
            IsDlt = True
            DltLin = CvActLin(DltDic(J)).Lin
        Else
            IsDlt = False
            DltLin = ""
        End If
        
        If InsDic.Exists(J) Then
            IsIns = True
            InsLin = CvActLin(InsDic(J)).Lin
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
FmtSrcMdy = O
End Function
Function SrcMdy(Src$(), B() As ActLin) As String()
ThwErMsg ErzActLinAy(B), CSub, "Error in ActLinAy", "ActMd Src", LyzActLinAy(B), Src
Dim O$(), J&
O = Src
For J = UB(B) To 0 Step -1
    O = SrcMdyLin(O, B(J))
Next
SrcMdy = O
End Function

Function MdMdy(A As CodeModule, B() As ActLin, Optional Silent As Boolean) As CodeModule
Dim NewLines$: NewLines = JnCrLf(SrcMdy(Src(A), B)): 'Brw NewLines: Stop
RplMd A, NewLines
End Function

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
            Case .Act = eeInsLin
                PushActEr O, "For same line, the Later one (CurLno) should be delete, but now it is insert", Ix, Cur, Las
            End Select
        Case Else
        End Select
    End Select
End With
ErzActLinCurLas = O
End Function

Private Function ErzActLinAy(A() As ActLin) As String()
Dim Ix%, Las As ActLin, Msg$, Cur As ActLin
If Si(A) <= 1 Then Exit Function
Set Las = A(0)
For Ix = 1 To UB(A)
    Set Cur = A(Ix)
    PushIAy ErzActLinAy, ErzActLinCurLas(Ix, Cur, Las) '<===
    Set Las = Cur
Next
End Function

Private Function LyzActLinAy(A() As ActLin) As String()
Dim J%
For J = 0 To UB(A)
    PushI LyzActLinAy, A(J).ToStr
Next
End Function
