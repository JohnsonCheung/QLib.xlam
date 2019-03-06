Attribute VB_Name = "MVb_Ay_Op_Srt"
Option Explicit

Function LinesSrt$(A$)
LinesSrt = JnCrLf(AySrt(SplitCrLf(A)))
End Function

Function IsSrtAy(A) As Boolean
Dim J&
For J = 0 To UB(A) - 1
   If A(J) > A(J + 1) Then Exit Function
Next
IsSrtAy = True
End Function

Function AyQSrt(A)
If Sz(A) = 0 Then Exit Function
Dim O: O = A
AyQSrtLH O, 0, UB(A)
AyQSrt = O
End Function

Sub AyQSrtLH(A, L&, H&)
If L >= H Then Exit Sub
Dim P&
P = AyQSrtPartition(A, L, H)
AyQSrtLH A, L, P
AyQSrtLH A, P + 1, H
End Sub

Function AyQSrtPartition&(A, L&, H&)
Dim V, I&, J&, X
V = A(L)
I = L - 1
J = H + 1
Dim Z&
Do
    Z = Z + 1
    If Z > 10000 Then Stop
    Do
        I = I + 1
    Loop Until A(I) >= V
    
    Do
        J = J - 1
    Loop Until A(J) <= V

    If I >= J Then
        AyQSrtPartition = J
        Exit Function
    End If

     X = A(I)
     A(I) = A(J)
     A(J) = X
Loop
End Function
Private Sub Z_AySrt_BY_AY()
Dim Ay, ByAy
Ay = Array(1, 2, 3, 4)
ByAy = Array(3, 4)
Ept = Array(3, 4, 1, 2)
GoSub Tst
Exit Sub
Tst:
    Act = AySrt_BY_AY(Ay, ByAy)
    C
    Return
End Sub

Function AySrt_BY_AY(Ay, ByAy)
Dim O: O = AyCln(Ay)
Dim I
For Each I In ByAy
    If HasEle(Ay, I) Then PushI O, I
Next
PushIAy O, AyMinus(Ay, O)
AySrt_BY_AY = O
End Function

Function AySrt(Ay, Optional Des As Boolean)
If Sz(Ay) = 0 Then AySrt = Ay: Exit Function
Dim Ix&, V, J&
Dim O: O = Ay: Erase O
Push O, Ay(0)
For J = 1 To UB(Ay)
    O = AyInsItm(O, Ay(J), AySrt__Ix(O, Ay(J), Des))
Next
AySrt = O
End Function

Private Function AySrt__Ix&(A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In A
        If V > I Then AySrt__Ix = O: Exit Function
        O = O + 1
    Next
    AySrt__Ix = O
    Exit Function
End If
For Each I In A
    If V < I Then AySrt__Ix = O: Exit Function
    O = O + 1
Next
AySrt__Ix = O
End Function

Function IxAyzAySrt(Ay, Optional Des As Boolean) As Long()
If Sz(Ay) = 0 Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = AyInsItm(O, J, IxAyzAySrt_Ix(O, Ay, Ay(J), Des))
Next
IxAyzAySrt = O
End Function

Private Sub Z_AySrt()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                   Act = AySrt(A):       ThwAyabNE Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(A, True): ThwAyabNE Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(A):       ThwAyabNE Exp, Act
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
Act = AySrt(A)
ThwAyabNE Exp, Act
End Sub

Private Function IxAyzAySrt_Ix&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then IxAyzAySrt_Ix& = O: Exit Function
        O = O + 1
    Next
    IxAyzAySrt_Ix& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then IxAyzAySrt_Ix& = O: Exit Function
    O = O + 1
Next
IxAyzAySrt_Ix& = O
End Function

Private Sub Z_IxAyzAySrt()
Dim A: A = Array("A", "B", "C", "D", "E")
ThwAyabNE Array(0, 1, 2, 3, 4), IxAyzAySrt(A)
ThwAyabNE Array(4, 3, 2, 1, 0), IxAyzAySrt(A, True)
End Sub

Private Function AySrtInToIxIxAy&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then AySrtInToIxIxAy& = O: Exit Function
        O = O + 1
    Next
    AySrtInToIxIxAy& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then AySrtInToIxIxAy& = O: Exit Function
    O = O + 1
Next
AySrtInToIxIxAy& = O
End Function


Function DicSrt(A As Dictionary) As Dictionary
If A.Count = 0 Then Set DicSrt = New Dictionary: Exit Function
Dim K
Set DicSrt = New Dictionary
For Each K In AyQSrt(A.Keys)
   DicSrt.Add K, A(K)
Next
End Function

Private Sub ZZ_AySrt()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                    Act = AySrt(A):        ThwIfNE Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(A, True): ThwIfNE Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(A):       ThwIfNE Exp, Act
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
Act = AySrt(A)
ThwIfNE Exp, Act
End Sub

Private Sub ZZ_IxAyzAySrt()
Dim A: A = Array("A", "B", "C", "D", "E")
ThwIfNE Array(0, 1, 2, 3, 4), IxAyzAySrt(A)
ThwIfNE Array(4, 3, 2, 1, 0), IxAyzAySrt(A, True)
End Sub


Private Sub Z()
Z_AySrt
Z_IxAyzAySrt
MVb__Srt:
End Sub