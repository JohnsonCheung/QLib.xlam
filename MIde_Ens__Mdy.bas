Attribute VB_Name = "MIde_Ens__Mdy"
Option Explicit
Sub MdyLin(A As CodeModule, B As ActLin, Optional Silent As Boolean)
Dim M As Boolean: M = Not Silent
Select Case B.Act
Case eActLin.eInsLin
    If M Then Inf CSub, "Line inserted", "Md Lin At", MdNm(A), B.Lin, B.Lno
    A.InsertLines B.Lno, B.Lin
Case eActLin.eDltLin
    If A.Lines(B.Lno, 1) <> B.Lin Then Stop
    If M Then Inf CSub, "Line deleted", "Md Lin At", MdNm(A), B.Lin, B.Lno
    A.DeleteLines B.Lno, 1
Case Else
    Thw CSub, "Invalid ActLin", "Md ActLin", MdNm(A), B
End Select
End Sub

Sub MdyPj(A As ActPj, Optional Silent As Boolean)
Dim B() As ActMd: B = A.ActMdAy
Dim J%
For J = 0 To UB(B)
    MdyMd B(J), Silent
Next
End Sub

Sub MdyMd(A As ActMd, Optional Silent As Boolean)
Dim J%, B() As ActLin: B = A.ActLinAy
ThwErMsg ActMdEr(A), CSub, "Error in ActMd", "ActMd Src", LyzActMd(A), AyAddIxPfx(Src(A), 1)
'BrwAy LyzActMd(B)
For J = UB(B) To 0 Step -1
    MdyLin A, B(J), Silent
Next
End Sub

Private Sub PushActEr(O$(), Msg$, Ix%, A() As ActLin)
Dim Nav(3)
Nav(0) = "Ix Las Cur"
Nav(1) = Ix
Nav(2) = A(Ix - 1).Lin
Nav(3) = A(Ix).ToStr
PushIAy O, LyzMsgNav(Msg, Nav) '<-----------
End Sub

Function ActMdEr(A As ActMd) As String()
Dim B() As ActLin
B = A.ActLinAy
Dim O$()
    Dim Ix%, Las As ActLin, Msg$, Cur As ActLin
    If Si(A) <= 1 Then Exit Function
    Set Las = A(0)
    For Ix = 1 To UB(A)
        Set Cur = A(Ix)
        With Cur
            Select Case True
            Case .Lno = 0
                PushActEr O, "Lno cannot be zero", Ix, B
            Case Not HasPfx(.Lin, "Const C")
                PushActEr O, "ActMd.Lin must with pfx Const C", Ix, B
            Case Else
                Select Case Las.Lno
                Case Is > .Lno: PushActEr O, "ActMd not in order", Ix, B
                Case .Lno:
                    Select Case True
                    Case .Act = Las.Act
                        PushActEr O, "Two same Lno should not same have 'IsIns'", Ix, B
                    Case .Act = eInsLin
                        PushActEr O, "For same line, the Later one (CurLno) should be delete, but now it is insert", Ix, B
                    End Select
                Case Else
                End Select
            End Select
        End With
        Set Las = A(Ix)
    Next
ActMdEr = O
End Function


Private Function LyzActMd(A As ActMd) As String()
Dim J%
For J = 0 To UB(A.ActLinAy)
    PushI LyzActMd, J & ":" & A(J).ToStr
Next
End Function

