Attribute VB_Name = "QVb_Lin_Lnx"
Option Explicit
Enum EmIxFm 'Ix is always 0.  When present as Lno, Ix=0 will be Lno=1.
    EiFm1   'Or, presenting Ix=0 as 1 for Lno
    EiFm0   'So, presenting Ix=0 as 0  for Ix
End Enum
Type SIxy
    S As String
    Ixy() As Long
End Type
Type SIxys: N As Integer: Ay() As SIxy: End Type
Type Lnx
    Lin As String
    Ix As Long
End Type
Type Lnxs: N As Long: Ay() As Lnx: End Type
Function Lnx(Lin$, Ix&) As Lnx
If Ix < 0 Then Thw CSub, "Ix cannot be negative", "Ix Lin", Ix, Lin
Lnx.Ix = Ix
Lnx.Lin = Lin
End Function
Sub PushLnx(O As Lnxs, M As Lnx)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function Lnxs(Ly$()) As Lnxs
Dim J&
For J = 0 To UB(Ly)
    PushLnx Lnxs, Lnx(Ly(J), J)
Next
End Function
Function IxywLin(A As Lnx, Lin$) As Long()
For J = 0 To A.N - 1
    With A.Ay(J)
    If .Lin = Lin Then PushI IxywLin, .Ix
    End With
Next
End Function

Function Lnoss$(Ixy() As Long)
Lnoss = JnSpc(AyIncEle1(Ixy))
End Function

Function LnosswLin$(A As Lnxs, Lin$)
LnosswLin = Lnoss(IxywLin(A, Lin))
End Function
Function LinSyzLnxs(A As Lnxs) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LinSyzLnxs, A.Ay(J).Lin
Next
End Function
Function SIxyszDup(A As Lnxs) As SIxys
Dim Dup$, I
For Each I In Itr(AywDup(LinSyzLnxs(A)))
    Dup = I
    PushSIxy SIxyszDup, SIxy(Dup, LnosszItm(A, Dup))
Next
End Function

Function BrwLnxs(A As Lnxs)
B LyzLnxs(A, EiFm1)
End Function
Sub BrwSIxys(A As SIxys)
B LyzSIxys(A)
End Sub
Function LyzSIxys(A As SIxys, Optional SLnossMacros$) As String()
Dim J&
For J = 0 To A.N - 1
    PushI LyzSIxys, LinzSIxy(A.Ay(J))
Next
End Function

Function LinzSIxy$(A As SIxy, Optional SLnossMacro$)
Dim M$: M = DftStr(SLnossMacro, "S({S}) Lnoss({Lnoss})")
With A
LinzSIxy = FmtMacro(M, .S, Lnoss(.Ixy))
End With
End Function
Sub PushSIxy(O As SIxys, M As SIxy)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Private Sub Z_DupT1zLnxs()
Dim A As Lnxs, Act As SIxys, Ept As SIxys
GoSub T0
'GoSub ZZ
Exit Sub
ZZ:
    A = Lnxs(SrcV)
    BrwSIxys DupT1zLnxs(A)
    Return
T0:
    A = Lnxs(Sy("A B", " B X", "A", "B", "C"))
    PushSIxy Ept, SIxy("A", Lngy(1, 3))
    PushSIxy Ept, SIxy("B", Lngy(2, 4))
    GoTo Tst
Tst:
    Act = DupT1zLnxs(A)
    BrwSIxys Act
    BrwSIxys Ept
    Debug.Assert IsEqSIxys(Act, Ept)
    Return
End Sub
Function IsEqSIxys(A As SIxys, B As SIxys) As Boolean
With A
If .N <> B.N Then Exit Function
Dim J&
For J = 0 To .N - 1
    If Not IsEqSIxy(.Ay(J), B.Ay(J)) Then Exit Function
Next
End With
IsEqSIxys = True
End Function
Function IsEqSIxy(A As SIxy, B As SIxy) As Boolean
With A
Select Case True
Case .Itm <> B.Itm, .Lnoss <> B.Lnoss
Case Else: IsEqSIxy = True
End Select
End With
End Function
Function DupT1zLnxs(A As Lnxs) As SIxys
DupT1zLnxs = DupLinzLnxs(T1Lnxs(A))
End Function
Function DupLinzLnxs(A As Lnxs) As SIxys
Dim Dup$, I
For Each I In Itr(AywDup(LinSyzLnxs(A)))
    Dup = I
    PushSIxy DupLinzLnxs, SIxy(Dup, IxyzLin(A, Dup))
Next
End Function
Function T1Lnxs(A As Lnxs) As Lnxs 'Take the T1 of A().Lin to return Lnxs
Const Insp As Boolean = True
If Insp Then BrwLnxs A
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
    PushLnx T1Lnxs, Lnx(T1(.Lin), .Ix)
    End With
Next
If Insp Then
    BrwLnxs T1Lnxs
    Stop
End If
End Function

Function LyzLnxs(A As Lnxs, Optional B As EmIxFm) As String()
Dim N0or1%: N0or1 = IIf(B = EiFm1, 1, 0)
Dim N%: N = MaxIx(A) + N0or1
Dim W%: W = Len(CStr(N))
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
    PushI LyzLnxs, AlignR(CStr(.Ix), W) & " " & .Lin
    End With
Next
End Function

Private Function MaxIx%(A As Lnxs)
Dim J&, O&
For J = 0 To A.N - 1
    O = Max(A.Ay(J).Ix, O)
Next
MaxIx = O
End Function

Function IxyzLin(A As Lnxs, Lin$) As Long()
Dim J&
For J = 0 To A.N - 1
    With A.Ay(J)
        If .Lin = Lin Then PushI IxyzLin, .Ix
    End With
Next
End Function

Function SIxy(S$, Ixy() As Long) As SIxy
SIxy.S = S
SIxy.Ixy = Ixy
End Function

Sub A()
Z_DupT1zLnxs
End Sub
