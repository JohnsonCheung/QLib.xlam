Attribute VB_Name = "QVb_Ay_Push"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Ay_Push."

Function Shf(OAy)
Shf = OAy(0)
OAy = AyeFstNEle(OAy)
End Function

Sub Push(O, M)
Dim N&
N = Si(O)
ReDim Preserve O(N)
If IsObject(M) Then
    Set O(N) = M
Else
    O(N) = M
End If
End Sub

Sub PushAp(O, ParamArray Ap())
Dim Av(), I: Av = Ap
For Each I In Av
    Push O, I
Next
End Sub

Sub PushAy(O, Ay)
Dim I
For Each I In Itr(Ay)
    Push O, I
Next
End Sub

Sub PushAyNoDup(O, Ay)
Dim I
For Each I In Itr(Ay)
    PushNoDup O, I
Next
End Sub

Sub PushI(O, M)
Dim N&
N = Si(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Sub PushIy(OSy$(), Sy$())
PushIAy OSy, Sy
End Sub

Sub PushIAy(O, MAy)
ThwIf_NotAy MAy, CSub
Dim M
For Each M In Itr(MAy)
    PushI O, M
Next
End Sub

Sub PushISomSi(OAy, IAy)
If Si(IAy) = 0 Then Exit Sub
PushI OAy, IAy
End Sub

Sub PushItmAy(O, Itm, Ay)
Push O, Itm
PushAy O, Ay
End Sub

Sub PushNoDup(O, M)
If Not HasEle(O, M) Then PushI O, M
End Sub
Sub PushNoDupNonBlankStr(O, M)
If M = "" Then Exit Sub
If Not HasEle(O, M) Then PushI O, M
End Sub
Sub PushNoDupDr(ODry(), Dr)
Dim IDr
For Each IDr In Itr(ODry)
    If IsEqAy(IDr, Dr) Then Exit Sub
Next
PushI ODry, Dr
End Sub
Sub PushNoDupAy(O, Ay)
Dim I
For Each I In Itr(Ay)
    PushNoDup O, I
Next
End Sub

Sub PushNonBlank(O$(), M)
If Trim(M) <> "" Then PushI O, M
End Sub

Sub PushNonBlankAy(O$(), MAy)
Dim I
For Each I In Itr(MAy)
    PushNonBlank O, I
Next
End Sub

Sub PushNonEmp(O, M)
If Not IsEmp(M) Then Push O, M
End Sub

Sub PushNonNothing(O, M)
If IsNothing(M) Then Exit Sub
PushObj O, M
End Sub

Sub PushNonZSz(O, Ay)
If Si(Ay) = 0 Then Exit Sub
PushI O, Ay
End Sub
Sub PushObjzExlNothing(O, M)
If IsNothing(M) Then PushObj O, M
End Sub

Sub PushObj(O, M)
Dim N&
N = Si(O)
ReDim Preserve O(N)
Set O(N) = M
End Sub

Sub PushObjzItr(O, Itr)
Dim Obj
For Each Obj In Itr
    PushObj O, Obj
Next
End Sub

Sub PushObjAy(O, Oy)
Dim I
For Each I In Itr(Oy)
    PushObj O, I
Next
End Sub

Sub PushWithSz(O, Ay)
If Not IsArray(Ay) Then Stop
If Si(Ay) = 0 Then Exit Sub
Push O, Ay
End Sub

Function Si&(A)
On Error Resume Next
Si = UBound(A) + 1
End Function

Function UB&(A)
UB = Si(A) - 1
End Function

Function Pop(O)
Asg LasEle(O), Pop
If Si(O) = 1 Then
    Erase O
Else
    ReDim Preserve O(UB(O) - 1)
End If
End Function
Function PopI(O)
PopI = LasEle(O)
If Si(O) = 1 Then
    Erase O
Else
    ReDim Preserve O(UB(O) - 1)
End If
End Function

Function AyReSzU(Ay, U&)
Dim O: O = Ay
If U < 0 Then
    Erase O
Else
    ReDim Preserve O(U)
End If
AyReSzU = O
End Function
