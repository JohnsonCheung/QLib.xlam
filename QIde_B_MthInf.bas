Attribute VB_Name = "QIde_B_MthInf"
Option Explicit
Option Compare Text
'*LnoC:Fun|Sub has one StartLineNo|Count.  Prp may have 2.
Type LnoC2: S1 As Long: C1 As Long: S2 As Long: C2 As Long: End Type
Type LnoC: S As Long: C As Long: End Type

Function MthLnoC(M As CodeModule, MthLno&) As LnoC
If MthLno = 0 Then Thw CSub, "MthLno cannot be zero"
With MthLnoC
    .S = MthLno
    If .S = 0 Then Exit Function
    Dim A&: A = EndLnozM(M, MthLno)
    .C = A - .S + 1: If .C <= 0 Then Thw CSub, FmtQQ("MthLineCnt[?] cannot be 0 or neg", .C)
End With
End Function

Function FmtLnoC2$(A As LnoC2)
With A
FmtLnoC2 = FmtQQ("LnoC(? ? ? ? ?)", .S1, .C1, "|", .S2, .C2)
End With
End Function

Function MthLnoC2(M As CodeModule, Mthn) As LnoC2
Dim A&(): A = MthLnoAyzMN(M, Mthn)
Dim O As LnoC2
Select Case Si(A)
Case 0
Case 1: GoSub X1
Case 2: GoSub X1: GoSub X2
Case Else: Thw CSub, "There is error in MthLnoAyzNM, it should return 0,1 or 2 Lno, but now[" & Si(A) & "]"
End Select
MthLnoC2 = O
Exit Function
X1:
    With MthLnoC(M, A(0))
    O.C1 = .C
    O.S1 = .S
    End With
    Return
X2:
    With MthLnoC(M, A(1))
    O.C2 = .C
    O.S2 = .S
    End With
    Return
End Function


'
