Attribute VB_Name = "QIde_B_MthInf"
Type LnoLines: Lno As Long: Lines As String: End Type
'*MthSC:Fun|Sub has one StartLineNo|Count.  Prp may have 2.
Type MthSC: S1 As Long: C1 As Long: S2 As Long: C2 As Long: End Type
Private Type SC: S As Long: C As Long: End Type

Function LnoLines(Lno&, Optional Lines$) As LnoLines
LnoLines.Lno = Lno
LnoLines.Lines = Lines
End Function
Private Function IsRmkzRmkBlk(RBlk$()) As Boolean
If Si(RBlk) = 0 Then Exit Function
Dim L$: L = LasEle(RBlk)
IsRmkzRmkBlk = Right(L, 2) = "@@"
End Function
Private Function RmkBlkzM(M As CodeModule, RLno&) As String()
If RLno = 0 Then Exit Function
Dim J&, L$, O$()
For J = RLno To M.CountOfLines
    L = M.Lines(J, 1)
    If Not IsVbRmk(L) Then Exit For
    PushI O, L
Next
RmkBlkzM = O
End Function

Private Function MthRLno&(M As CodeModule, MthLno&)
Dim ELin$: ELin = EndLinzM(M, MthLno)
Dim Lno&: For Lno = MthLno + 1 To M.CountOfLines
    Dim L$: L = M.Lines(Lno, 1)
    If IsVbRmk(L) Then MthRLno = Lno: Exit Function
    If L = ELin Then Exit Function
Next
ThwImpossible CSub
End Function
Function MthRmkL$(M As CodeModule, MthLno&)
MthRmkL = MthRmkLzRLno(M, MthRLno(M, MthLno))
End Function
Function MthRmkLzRLno$(M As CodeModule, RLno&)
Dim RBlk$(): RBlk = RmkBlkzM(M, RLno)
Dim IsRmk As Boolean: IsRmk = IsRmkzRmkBlk(RBlk)
If IsRmk Then MthRmkLzRLno = JnCrLf(RBlk)
End Function

Function MthRmkzL(M As CodeModule, MthLno&) As LnoLines
'Ret MthRmkzL : Lno Lines ! MthRmk is fst gp of rmk lines and its last lin las 2 char is
'                         ! if no rmk is fnd, the Lno is the NxtLno-of-Md-MthLno. @@
If MthLno = 0 Then Exit Function
Dim RLno&: RLno = MthRLno(M, MthLno)
Dim RmkL$: RmkL = MthRmkLzRLno(M, RLno)
If RmkL = "" Then
    Dim Nxt&:  Nxt = NxtLnozML(M, MthLno)
    MthRmkzL = LnoLines(Nxt)
Else
    MthRmkzL = LnoLines(RLno, RmkL)
End If
End Function

Private Sub Z_MthRmkzM()
Dim M As CodeModule: Set M = Md("QIde_B_MthOp__AlignMthDimzML")
BrwS1S2s MthRmkzM(M)
End Sub

Function MthRmkzM(M As CodeModule) As S1S2s
Dim S$(): S = Src(M)
Dim Ix, O As S1S2s: For Each Ix In MthIxItr(S)
    Dim Lno&: Lno = Ix + 1
    Dim R$: R = MthRmkL(M, Lno)
    If R <> "" Then
        Dim N$: N = Mthn(S(Ix))
        PushS1S2 O, S1S2(N, R)
    End If
Next
MthRmkzM = O
End Function

Function MthRmk(M As CodeModule, Mthn) As LnoLines
MthRmk = MthRmkzL(M, MthLnozMM(M, Mthn))
End Function

Function MthSC1(M As CodeModule, MthLno&) As SC
If MthLno = 0 Then Thw CSub, "MthLno cannot be zero"
With MthSC1
    .S = MthLno
    If .S = 0 Then Exit Function
    Dim A&: A = EndLnozM(M, MthLno)
    .C = A - .S + 1: If .C <= 0 Then Thw CSub, FmtQQ("MthLineCnt[?] cannot be 0 or neg", .C)
End With
End Function
Function LinzMthSC$(A As MthSC)
With A
LinzMthSC = FmtQQ("MthSC(? ? ? ? ?)", .S1, .C1, "|", .S2, .C2)
End With
End Function
Function MthSC(M As CodeModule, Mthn) As MthSC
Dim A&(): A = MthLnoAyzMN(M, Mthn)
Dim O As MthSC
Select Case Si(A)
Case 0
Case 1: GoSub X1
Case 2: GoSub X1: GoSub X2
Case Else: Thw CSub, "There is error in MthLnoAyzNM, it should return 0,1 or 2 Lno, but now[" & Si(A) & "]"
End Select
MthSC = O
Exit Function
X1:
    With MthSC1(M, A(0))
    O.C1 = .C
    O.S1 = .S
    End With
    Return
X2:
    With MthSC1(M, A(1))
    O.C2 = .C
    O.S2 = .S
    End With
    Return
End Function
