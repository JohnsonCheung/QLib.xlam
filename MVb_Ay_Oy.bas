Attribute VB_Name = "MVb_Ay_Oy"
Option Explicit
Function OyAdd(Oy1, Oy2)
Dim O: O = Oy1
PushObjAy O, Oy2
OyAdd = O
End Function

Sub DoItoMth(Ito, ObjMth$)
Dim Obj As Object
For Each Obj In Ito
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Sub DoOyMth(Oy() As Object, ObjMth$)
Dim Obj As Object
For Each Obj In Itr(Oy)
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Function FstItmzOyPEv(Oy, P$, Ev)
Dim Obj As Object
For Each Obj In Itr(Oy)
    If Prp(Obj, P) = Ev Then Asg Obj, FstItmzOyPEv: Exit Function
Next
End Function

Function AvzOyP(Oy, P$) As Variant()
AvzOyP = IntozOyP(EmpAv, Oy, P)
End Function

Function IntozOyP(Into, Oy, P$)
Dim O: O = Into: Erase O
Dim Obj As Object
For Each Obj In Itr(Oy)
    Push O, Prp(Obj, P)
Next
IntozOyP = O
End Function

Function IntAyzOyP(Oy, P$) As Integer()
IntAyzOyP = IntozOyP(EmpIntAy, Oy, P)
End Function

Function SyzOyPrp(Oy, P$) As String()
SyzOyPrp = IntozOyP(EmpSy, Oy, P)
End Function

Function OyRmvFstNEle(Oy, N&)
Dim O: O = Oy
ReDim O(N - 1)
Dim J&
For J = 0 To UB(Oy) - N
    Set O(J) = Oy(N + J)
Next
OyRmvFstNEle = O
End Function

Function OyeNothing(Oy)
OyeNothing = AyCln(Oy)
Dim Obj As Object
For Each Obj In Oy
    If Not IsNothing(Obj) Then PushObj OyeNothing, Obj
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

Function OywPredXPTrue(Oy, XP$, P$)
Dim O, Obj As Object
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If Run(XP, Obj, P) Then
        PushObj O, Obj
    End If
Next
OywPredXPTrue = O
End Function

Function OywPEv(Oy, P$, Ev)
Dim O
   O = Oy
   Erase O
   Dim Obj As Object
   For Each Obj In Itr(Oy)
       If Prp(Obj, P) = Ev Then PushObj O, Obj
   Next
OywPEv = O
End Function

Function OywPInAy(Oy, P$, InAy)
Dim Obj As Object, O
If Si(Oy) = 0 Or Si(InAy) Then OywPInAy = Oy: Exit Function
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If HasEle(InAy, Prp(Obj, P)) Then PushObj O, Obj
Next
OywPInAy = O
End Function
Function LyzObjPP(Obj As Object, PP$) As String()
Dim I
For Each I In SySsl(PP)
    PushS LyzObjPP, I & " " & Prp(Obj, CStr(I))
Next
End Function
Function DryzOyPP(Oy, PP$) As Variant()
Dim Obj As Object, PrpPthSy$()
PrpPthSy = SySsl(PP)
For Each Obj In Itr(Oy)
    PushI DryzOyPP, DrzObj(Obj, PrpPthSy)
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
