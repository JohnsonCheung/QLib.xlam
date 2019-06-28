Attribute VB_Name = "QVb_Ay_Op_Srt"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Srt."
Private Const Asm$ = "QVb"
Enum EmOrd
    EiAsc
    EiDes
End Enum
Private A_Ay
Function SrtLines$(A$)
SrtLines = JnCrLf(SrtAy(SplitCrLf(A)))
End Function

Function IsSrtedzAy(Ay) As Boolean
Dim J&
For J = 0 To UB(Ay) - 1
   If Ay(J) > Ay(J + 1) Then Exit Function
Next
IsSrtedzAy = True
End Function

Private Sub Z_SrtAyQ()
Dim Ay, IsDes As Boolean
GoSub T0
GoSub T1
Exit Sub
T0:
    Ay = Array(1, 2, 3, 4, 0, 1, 1, 5)
    IsDes = False
    Ept = Array(0, 1, 1, 1, 2, 3, 4, 5)
    GoTo Tst
T1:
    Ay = Array(1, 2, 4, 87, 4, 2)
    IsDes = True
    Ept = Array(87, 4, 4, 2, 2, 1)
    GoTo Tst
Tst:
    Act = SrtAyQ(Ay, IsDes:=IsDes)
    C
    Return
End Sub


Function AwEQ(Ay, V)
If Si(Ay) <= 1 Then AwEQ = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I = V Then PushI O, I
Next
AwEQ = O
End Function

Function AwLE(Ay, V)
If Si(Ay) <= 1 Then AwLE = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I <= V Then PushI O, I
Next
AwLE = O
End Function
Function AwLT(Ay, V)
If Si(Ay) = 1 Then AwLT = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I < V Then PushI O, I
Next
AwLT = O
End Function
Function AwGT(Ay, V)
If Si(Ay) = 1 Then AwGT = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I > V Then PushI O, I
Next
AwGT = O
End Function

Function SrtAyQ(Ay, Optional IsDes As Boolean)
If Si(Ay) = 0 Then SrtAyQ = Ay: Exit Function
A_Ay = Ay  ' Put it is A_Ay, which is untouch, only sort the %Ixy
Dim Ixy&(): Ixy = LngSeqzFT(0, UB(Ay))
Ixy = SrtAyQ__Srt(Ixy)  'Srt the %Ixy
If IsDes Then
    Ixy = RevAyI(Ixy)
End If
SrtAyQ = AwIxy(Ay, Ixy)
End Function

Private Function SrtAyQ__Srt(Ixy&()) As Long()
'Ret : a sorted ix of @Ixy
Dim O&()
    Select Case Si(Ixy)
    Case 0
    Case 1: O = Ixy
    Case 2:
        If SrtAyQ__IsLE(Ixy(0), Ixy(1)) Then
            O = Ixy
        Else
            PushI O, Ixy(1)
            PushI O, Ixy(0)
        End If
    Case Else
        Dim L&():   L = Ixy
        Dim P&:   P = Pop(L)
        Dim A&(): A = SrtAyQ__LE(L, P)
        Dim B&(): B = SrtAyQ__GT(L, P)
                  A = SrtAyQ__Srt(A) 'Srt-it
                  B = SrtAyQ__Srt(B)
        PushIAy O, A
          PushI O, P
        PushIAy O, B
    End Select
SrtAyQ__Srt = O
End Function
Private Function SrtAyQ__IsLE(A, B&) As Boolean
SrtAyQ__IsLE = A_Ay(A) <= A_Ay(B)
End Function
Private Function SrtAyQ__LE(Ixy&(), P&) As Long()
'@Ixy : Ix to be selected
'@P   : the Pivot-Ix to select those ix in @Ixy
'Ret : a subset of @Ixy so that each the ret ix is LE than the Pivot-@P
Dim I: For Each I In Itr(Ixy)
    If SrtAyQ__IsLE(I, P) Then PushI SrtAyQ__LE, I  ' If the running-Ix-%I IsGT pivot-Ix-@P, push it to ret
Next
End Function
Private Function SrtAyQ__GT(Ixy&(), P&) As Long()
'@Ixy : Ix to be selected
'@P   : the Pivot-Ix to select those ix in @Ixy
'Ret : a subset of @Ixy so that each the ret ix is GT than the Pivot-@P
Dim I: For Each I In Itr(Ixy)
    If Not SrtAyQ__IsLE(I, P) Then PushI SrtAyQ__GT, I  ' If the running-Ix-%I IsGT pivot-Ix-@P, push it to ret
Next
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
Dim O: O = ResiU(Ay)
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
    O = InsEle(O, Ay(J), SrtAy__Ix(O, Ay(J)))
Next
If Des Then
    SrtAy = RevAy(O)
Else
    SrtAy = O
End If
End Function

Private Function SrtAy__Ix&(A, V)
Dim I, O&
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
    O = InsEle(O, J, IxyzSrtAy_Ix(O, Ay, Ay(J), Des))
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
Act = SrtAyQ(A)
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

Function DiczAddIxToKey(A As Dictionary) As Dictionary
Dim O As New Dictionary, K, J&
For Each K In A.Keys
    O.Add J & " " & K, A(K)
    J = J + 1
Next
Set DiczAddIxToKey = O
End Function

Function SrtDic(A As Dictionary, Optional IsDesc As Boolean) As Dictionary
If A.Count = 0 Then Set SrtDic = New Dictionary: Exit Function
Dim O As New Dictionary
Dim Srt: Srt = SrtAyQ(A.Keys, IsDesc)
Dim K: For Each K In Srt
   O.Add K, A(K)
Next
Set SrtDic = O
End Function

Private Sub Z_SrtAy4()
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

Private Sub Z_IxyzSrtAy5()
Dim A: A = Array("A", "B", "C", "D", "E")
ThwIf_NE Array(0, 1, 2, 3, 4), IxyzSrtAy(A)
ThwIf_NE Array(4, 3, 2, 1, 0), IxyzSrtAy(A, True)
End Sub


Private Sub Z()
Z_SrtAy
Z_IxyzSrtAy
MVb__Srt:
End Sub
