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
    SrcMdyLin = CvSy(AyInsItm(Src, B.Lin, B.Lno))
Case eActLin.eeDltLin
    If Src(B.Lno - 1) <> B.Lin Then Stop
    SrcMdyLin = AyeEleAt(Src, B.Lno - 1)
Case Else
    Thw CSub, "Invalid ActLin.Act, it should eeInsLin or eeDltLin, ", "ActLin", B.ToStr
End Select
End Function

Function PjMdy(A As VBProject, B() As ActMd, Optional Silent As Boolean) As VBProject
Dim I
For Each I In Itr(B)
    With CvActMd(I)
        MdMdy .Md, .ActLinAy, Silent
    End With
Next
Set PjMdy = A
End Function
Sub Z()

End Sub
Function LyzSrcApplyMdy(Src$(), B() As ActLin) As String()
Dim N As Byte: N = NDig(Si(Src))
Dim J&, Middle$, Lno$, Lin$, O$()
Dim D As New Dictionary
    For J = 0 To UB(B)
        D.Add B(J).Ix, B(J)
    Next
For J = 0 To UB(Src)
    Dim Hit As Boolean: D.Exists (J)
    Dim M As ActLin: If Hit Then Set M = D(J) Else Set M = Nothing
    Lin = Src(J)
    Middle = IIf(Hit, " <<<<< ", "        ")
    Lno = AlignL(J + 1, N)
    PushI O, Lno & Middle & Lin
    If D.Exists(J) Then
        PushI O, Lno & " >>>>> " & M.Lin
    End If
Next
LyzSrcApplyMdy = O
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
Stop
MdRpl A, NewLines
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
