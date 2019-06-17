Attribute VB_Name = "QIde_B_MthInf"
Option Explicit
Option Compare Text
'*MthSC:Fun|Sub has one StartLineNo|Count.  Prp may have 2.
Type MthSC: S1 As Long: C1 As Long: S2 As Long: C2 As Long: End Type
Private Type SC: S As Long: C As Long: End Type

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

