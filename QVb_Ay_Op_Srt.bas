Attribute VB_Name = "QVb_Ay_Op_Srt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Srt."
Private Const Asm$ = "QVb"
Enum EmOrd
    EiAsc
    EiDes
End Enum
Function SrtLines$(A$)
SrtLines = JnCrLf(SrtAy(SplitCrLf(A)))
End Function

Function IsSrtedAy(Ay) As Boolean
Dim J&
For J = 0 To UB(Ay) - 1
   If Ay(J) > Ay(J + 1) Then Exit Function
Next
IsSrtedAy = True
End Function

Private Sub Z_QSrt()
Dim Act, Ay
GoSub T0
Exit Sub
T0:
    Ay = Array(1, 2, 4, 87, 4, 2)
    Ept = Array(87, 4, 4, 2, 2, 1)
    GoTo Tst
Tst:
    Act = QSrt(Ay, EiDes)
    Stop
    C
    Return
End Sub

Function QSrt(Ay, Optional Ord As EmOrd)
Dim N&: N = Si(Ay)
If N <= 1 Then QSrt = Ay: Exit Function
Dim V
    V = Ay(0)
Dim L, H
    L = AywLT(Ay, V):
    H = AywGT(Ay, V):
Dim L1, V1, H1
    L1 = QSrt(L)
    V1 = AywEQ(Ay, V)
    H1 = QSrt(H)
Dim O
    O = AddAyAp(L1, V1, H1)
If Ord = EiDes Then O = Reverse(O)
QSrt = O
End Function

Function AywEQ(Ay, V)
If Si(Ay) <= 1 Then AywEQ = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I = V Then PushI O, I
Next
AywEQ = O
End Function

Function AywLE(Ay, V)
If Si(Ay) <= 1 Then AywLE = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I <= V Then PushI O, I
Next
AywLE = O
End Function
Function AywLT(Ay, V)
If Si(Ay) = 1 Then AywLT = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I < V Then PushI O, I
Next
AywLT = O
End Function
Function AywGT(Ay, V)
If Si(Ay) = 1 Then AywGT = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I > V Then PushI O, I
Next
AywGT = O
End Function
Function QSrt1(Ay, Optional IsDes As Boolean)
If Si(Ay) = 0 Then Exit Function
Dim O: O = Ay
QSrt1LH O, 0, UB(Ay)
If IsDes Then
    QSrt1 = ReverseAyI(O)
Else
    QSrt1 = O
End If
End Function
Private Sub QSrt1LH(Ay, L&, H&)
If L >= H Then Exit Sub
Dim P&
P = QSrt1Partition(Ay, L, H)
QSrt1LH Ay, L, P
QSrt1LH Ay, P + 1, H
End Sub
Function QSrt1Partition1&(OAy, L&, H&) 'Try mdy
Dim V, I&, J&, X
V = OAy(L)
I = L
J = H
Dim Z&
Z = 0
Do
    Z = Z + 1
    If Z > 10000 Then Stop
    While OAy(I) < V
        I = I + 1
    Wend
    
    While OAy(J) > V
        J = J - 1
    Wend
    If I >= J Then
        QSrt1Partition1 = J
        Exit Function
    End If

    X = OAy(I)
    OAy(I) = OAy(J)
    OAy(J) = X
Loop
End Function

Function QSrt1Partition&(OAy, L&, H&)
Dim V, I&, J&, X
V = OAy(L)
I = L - 1
J = H + 1
Dim Z&
Do
    Z = Z + 1
    If Z > 10000 Then Stop
    Do
        I = I + 1
    Loop Until OAy(I) >= V
    
    Do
        J = J - 1
    Loop Until OAy(J) <= V

    If I >= J Then
        QSrt1Partition = J
        Exit Function
    End If

    X = OAy(I)
    OAy(I) = OAy(J)
    OAy(J) = X
Loop
End Function

Private Sub Z_SrtAyByAy()
Dim Ay, ByAy
Ay = Array(1, 2, 3, 4)
ByAy = Array(3, 4)
Ept = Array(3, 4, 1, 2)
GoSub Tst
Exit Sub
Tst:
    Act = SrtAyByAy(Ay, ByAy)
    C
    Return
End Sub

Function SrtAyByAy(Ay, ByAy)
Dim O: O = Resi(Ay)
Dim I
For Each I In ByAy
    If HasEle(Ay, I) Then PushI O, I
Next
PushIAy O, MinusAy(Ay, O)
SrtAyByAy = O
End Function

Function SrtAy(Ay, Optional Des As Boolean)
If Si(Ay) = 0 Then SrtAy = Ay: Exit Function
Dim Ix&, V, J&
Dim O: O = Ay: Erase O
Push O, Ay(0)
For J = 1 To UB(Ay)
    O = AyInsEle(O, Ay(J), SrtAy__Ix(O, Ay(J), Des))
Next
SrtAy = O
End Function

Private Function SrtAy__Ix&(A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In A
        If V > I Then SrtAy__Ix = O: Exit Function
        O = O + 1
    Next
    SrtAy__Ix = O
    Exit Function
End If
For Each I In A
    If V < I Then SrtAy__Ix = O: Exit Function
    O = O + 1
Next
SrtAy__Ix = O
End Function

Function IxyzSrtAy(Ay, Optional Des As Boolean) As Long()
If Si(Ay) = 0 Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = AyInsEle(O, J, IxyzSrtAy_Ix(O, Ay, Ay(J), Des))
Next
IxyzSrtAy = O
End Function

Private Sub Z_SrtAy()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                   Act = SrtAy(A):       ThwIf_AyabNE Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = SrtAy(A, True): ThwIf_AyabNE Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = SrtAy(A):       ThwIf_AyabNE Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin:Sub"
Push A, "Private:HasPfx:Function"
Push A, "Private:MdMthDRsFunBdyLy:Function"
Push A, "Private:SrcMthLx_ToLx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:HasPfx:Function"
Push Exp, "Private:MdMthDRsFunBdyLy:Function"
Push Exp, "Private:SrcMthLx_ToLx:Function"
Push Exp, "~Private:JnContinueLin:Sub"
Push Exp, "~~:Tst:Sub"
Act = QSrt1(A)
ThwIf_AyabNE Exp, Act
End Sub

Private Function IxyzSrtAy_Ix&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then IxyzSrtAy_Ix& = O: Exit Function
        O = O + 1
    Next
    IxyzSrtAy_Ix& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then IxyzSrtAy_Ix& = O: Exit Function
    O = O + 1
Next
IxyzSrtAy_Ix& = O
End Function

Private Sub Z_IxyzSrtAy()
Dim A: A = Array("A", "B", "C", "D", "E")
ThwIf_AyabNE Array(0, 1, 2, 3, 4), IxyzSrtAy(A)
ThwIf_AyabNE Array(4, 3, 2, 1, 0), IxyzSrtAy(A, True)
End Sub

Private Function SrtAyInEIxIxy&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then SrtAyInEIxIxy& = O: Exit Function
        O = O + 1
    Next
    SrtAyInEIxIxy& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then SrtAyInEIxIxy& = O: Exit Function
    O = O + 1
Next
SrtAyInEIxIxy& = O
End Function

Function DicAddIxToKey(A As Dictionary) As Dictionary
Dim O As New Dictionary, K, J&
For Each K In A.Keys
    O.Add J & " " & K, A(K)
    J = J + 1
Next
Set DicAddIxToKey = O
End Function

Function SrtDic(A As Dictionary, Optional IsDesc As Boolean) As Dictionary
If A.Count = 0 Then Set SrtDic = New Dictionary: Exit Function
Dim K
Set SrtDic = New Dictionary
For Each K In QSrt1(A.Keys, IsDesc)
   SrtDic.Add K, A(K)
Next
End Function

Private Sub ZZ_SrtAy()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                    Act = SrtAy(A):        ThwIf_NE Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = SrtAy(A, True): ThwIf_NE Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = SrtAy(A):       ThwIf_NE Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin:Sub"
Push A, "Private:HasPfx:Function"
Push A, "Private:MdMthDRsFunBdyLy:Function"
Push A, "Private:SrcMthLx_ToLx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:HasPfx:Function"
Push Exp, "Private:MdMthDRsFunBdyLy:Function"
Push Exp, "Private:SrcMthLx_ToLx:Function"
Push Exp, "~Private:JnContinueLin:Sub"
Push Exp, "~~:Tst:Sub"
Act = SrtAy(A)
ThwIf_NE Exp, Act
End Sub

Private Sub ZZ_IxyzSrtAy()
Dim A: A = Array("A", "B", "C", "D", "E")
ThwIf_NE Array(0, 1, 2, 3, 4), IxyzSrtAy(A)
ThwIf_NE Array(4, 3, 2, 1, 0), IxyzSrtAy(A, True)
End Sub


Private Sub ZZ()
Z_SrtAy
Z_IxyzSrtAy
MVb__Srt:
End Sub
