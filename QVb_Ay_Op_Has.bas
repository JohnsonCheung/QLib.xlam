Attribute VB_Name = "QVb_Ay_Op_Has"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Has."
Private Const Asm$ = "QVb"
Function HasObj(ObjAy, Obj) As Boolean
Dim OPtr&: OPtr = ObjPtr(Obj)
Dim I: For Each I In ObjAy
    If ObjPtr(I) = OPtr Then HasObj = True: Exit Function
Next
End Function
Function HasDup(Ay) As Boolean
Dim S As New Dictionary
Dim I: For Each I In Ay
    If S.Exists(I) Then HasDup = True: Exit Function
    S.Add I, Empty
Next
End Function

Function HasEleS(Ay, StrEle$, Optional C As VbCompareMethod = vbBinaryCompare) As Boolean
'Ret: true if @Ay has @StrEle
Dim I: For Each I In Itr(Ay)
    If IsEqStr(I, StrEle, C) Then HasEleS = True: Exit Function
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

Function IsAySub(SubAy, SuperAy) As Boolean
Dim I: For Each I In Itr(SubAy)
    If Not HasEle(SuperAy, I) Then Exit Function
Next
IsAySub = True
End Function

Function IsAySuper(SuperAy, SubAy) As Boolean
IsAySuper = IsAySub(SubAy, SuperAy)
End Function
Function ThwNotSuperAy(SuperAy, SubAy) As String()
If IsAySuper(SuperAy, SubAy) Then Exit Function
Thw CSub, "Some element in SubAy are found in SuperAy", "[Som Ele in SubAy] SubAy SuperAy", MinusAy(SubAy, SuperAy), SubAy, SuperAy
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

Function HasElePredXPTrue(A, Xp$, P) As Boolean
If Si(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(Xp, X, P) Then
        HasElePredXPTrue = True
        Exit Function
    End If
Next
End Function

Private Sub Z_HasEleAyInSeq()
Dim A, B
A = Array(1, 2, 3, 4, 5, 6, 7, 8)
B = Array(2, 4, 6)
Debug.Assert HasEleAyInSeq(A, B) = True

End Sub

Private Sub Z_HasEleDupEle()
Ass HasEleDupEle(Array(1, 2, 3, 4)) = False
Ass HasEleDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub
