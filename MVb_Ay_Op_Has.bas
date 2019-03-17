Attribute VB_Name = "MVb_Ay_Op_Has"
Option Explicit
Function HasObj(Ay, Obj) As Boolean
Dim I, OPtr&
OPtr = ObjPtr(Obj)
For Each I In Ay
    If ObjPtr(I) = OPtr Then HasObj = True: Exit Function
Next
End Function

Function HasEle(Ay, Ele) As Boolean
Dim I
For Each I In Itr(Ay)
    If I = Ele Then HasEle = True: Exit Function
Next
End Function

Function HasEleAy(Ay, EleAy) As Boolean
Dim I
For Each I In Itr(EleAy)
    If Not HasEle(Ay, I) Then Exit Function
Next
HasEleAy = True
End Function

Function HasElezInSomAyzOfAp(ParamArray AyAp()) As Boolean
Dim AvAp(): AvAp = AyAp
Dim Ay
For Each Ay In Itr(AvAp)
    If Si(Ay) > 0 Then HasElezInSomAyzOfAp = True: Exit Function
Next
End Function
Function IsSubAy(SubAy, SuperAy) As Boolean
Dim I
For Each I In Itr(SubAy)
    If Not HasEle(SuperAy, I) Then Exit Function
Next
IsSubAy = True
End Function
Function IsSuperAy(SuperAy, SubAy) As Boolean
IsSuperAy = IsSubAy(SubAy, SuperAy)
End Function
Function ThwNotSuperAy(SuperAy, SubAy) As String()
If IsSuperAy(SuperAy, SubAy) Then Exit Function
Thw CSub, "Some element in SubAy are found in SuperAy", "[Som Ele in SubAy] SubAy SuperAy", AyMinus(SubAy, SuperAy), SubAy, SuperAy
End Function

Function HasEleAyInSeq(A, B) As Boolean
Dim BItm, Ix&
If Si(B) = 0 Then Stop
For Each BItm In B
    Ix = IxzAy(A, BItm, Ix)
    If Ix = -1 Then Exit Function
    Ix = Ix + 1
Next
HasEleAyInSeq = True
End Function

Function HasEleDupEle(A) As Boolean
If Si(A) = 0 Then Exit Function
Dim Pool: Pool = A: Erase Pool
Dim I
For Each I In A
    If HasEle(Pool, I) Then HasEleDupEle = True: Exit Function
    Push Pool, I
Next
End Function

Function HasEleNegOne(A) As Boolean
Dim V
If Si(A) = 0 Then Exit Function
For Each V In A
    If V = -1 Then HasEleNegOne = True: Exit Function
Next
End Function

Function HasElePredPXTrue(A, PX$, P) As Boolean
Dim X
For Each X In Itr(A)
    If Run(PX, P, X) Then HasElePredPXTrue = True: Exit Function
Next
End Function

Function HasElePredXPTrue(A, XP$, P) As Boolean
If Si(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(XP, X, P) Then
        HasElePredXPTrue = True
        Exit Function
    End If
Next
End Function

Function IsAySub(Ay, SubAy) As Boolean
If Si(Ay) = 0 Then Exit Function
If Si(SubAy) = 0 Then IsAySub = True: Exit Function
Dim I
For Each I In SubAy
    If Not HasEle(Ay, I) Then Exit Function
Next
IsAySub = True
End Function

Private Sub ZZ_HasEleAyInSeq()
Dim A, B
A = Array(1, 2, 3, 4, 5, 6, 7, 8)
B = Array(2, 4, 6)
Debug.Assert HasEleAyInSeq(A, B) = True

End Sub

Private Sub ZZ_HasEleDupEle()
Ass HasEleDupEle(Array(1, 2, 3, 4)) = False
Ass HasEleDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub
