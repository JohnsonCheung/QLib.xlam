Attribute VB_Name = "MVb_Ay_Oy"
Option Explicit
Function OyAdd(A, B)
Dim O, I
O = A
For Each I In Itr(B)
    PushObj O, I
Next
OyAdd = O
End Function

Sub DoItrMth(Itr, ObjMth$)
Dim Obj
For Each Obj In Itr
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Sub DoOyMth(Oy, ObjMth$)
Dim Obj
For Each Obj In Itr(Oy)
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Function FstOyPEv(Oy, P, V)
Dim Obj
For Each Obj In Itr(Oy)
    If ObjPrp(Obj, P) = V Then Asg Obj, FstOyPEv: Exit Function
Next
End Function

Function AvOyP(Oy, P) As Variant()
AvOyP = IntoOyP(EmpAv, Oy, P)
End Function

Function IntoOyP(Into, Oy, P)
Dim O: O = Into: Erase O
Dim Obj
For Each Obj In Itr(Oy)
    Push O, ObjPrp(Obj, P)
Next
IntoOyP = O
End Function

Function IntAyOyP(A, P) As Integer()
IntAyOyP = IntoOyP(A, P, EmpIntAy)
End Function

Function SyzOyPrp(A, P) As String()
SyzOyPrp = IntoOyP(EmpSy, A, P)
End Function

Function OyRmvFstNEle(A, N&)
Dim O: O = A
ReDim O(N - 1)
Dim J&
For J = 0 To UB(A) - N
    Set O(J) = A(N + J)
Next
OyRmvFstNEle = O
End Function

Function OyeNothing(A)
OyeNothing = AyCln(A)
Dim I
For Each I In A
    If Not IsNothing(I) Then PushObj OyeNothing, I
Next
End Function

Function OywNmPfx(Oy, NmPfx$)
Dim Obj, O
O = Oy: Erase O
For Each Obj In Itr(Oy)
    If HasPfx(Obj.Name, NmPfx) Then PushObj O, Obj
Next
OywNmPfx = O
End Function

Function OywNm(Oy, B As WhNm)
Dim Obj, O
O = Oy: Erase O
For Each Obj In Itr(Oy)
    If HitAy(Obj.Name, B) Then PushObj OywNm, Obj
Next
End Function

Function OywPredXPTrue(A, XP$, P)
Dim O, X
O = A
Erase O
For Each X In Itr(A)
    If Run(XP, X, P) Then
        PushObj A, X
    End If
Next
OywPredXPTrue = O
End Function

Function OywPEv(Oy, P, Ev)
Dim O
   O = Oy
   Erase O
   Dim Obj
   For Each Obj In Itr(Oy)
       If ObjPrp(Obj, P) = Ev Then PushObj O, Obj
   Next
OywPEv = O
End Function

Function IntAyOywPEvSelP(Oy, P, Ev, SelP) As Integer()
IntAyOywPEvSelP = IntAyOyP(OywPEv(Oy, P, Ev), SelP)
End Function

Function DryOywPEvSelPP(Oy, P, Ev, SelPP$) As Variant()
DryOywPEvSelPP = DryOySelPP(OywPEv(Oy, P, Ev), SelPP)
End Function

Function OywPIn(A, P, InAy)
Dim X, O
If Si(A) = 0 Or Si(InAy) Then OywPIn = A: Exit Function
O = A
Erase O
For Each X In Itr(A)
    If HasEle(InAy, ObjPrp(X, P)) Then PushObj O, X
Next
OywPIn = O
End Function

Function DryOySelPP(Oy, SelPP$) As Variant()
Dim Obj, PrpNy$()
PrpNy = SySsl(SelPP)
For Each Obj In Itr(Oy)
    PushI DryOySelPP, DrzObjPrpNy(Obj, PrpNy)
Next
End Function

Private Sub ZZ_OyDrs()
'WsVis DrsNewWs(OyDrs(CurrentDb.TableDefs("ZZ_UpdSeqFld").Fields, "Name Type OrdinalPosition"))
End Sub

Private Sub ZZ_OyP_Ay()
Dim CdPanAy() As CodePane
Stop
'CdPanAy = Oy(CurPj.MdAy).PrpAy("CodePane", CdPanAy)
Stop
End Sub
