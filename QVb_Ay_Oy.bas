Attribute VB_Name = "QVb_Ay_Oy"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Oy."
Private Const Asm$ = "QVb"
Function OyAdd(Oy1, Oy2)
Dim O: O = Oy1
PushObjAy O, Oy2
OyAdd = O
End Function

Sub DoItoMth(ITo, ObjMth)
Dim Obj As Object
For Each Obj In ITo
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Sub DoOyMth(Oy, ObjMth)
Dim Obj
For Each Obj In Itr(Oy)
    CallByName Obj, ObjMth, VbMethod
Next
End Sub

Function FstItmzOyPEv(Oy, P As PrpPth, Ev)
Dim Obj As Object
For Each Obj In Itr(Oy)
    If Prp(Obj, P) = Ev Then Asg Obj, FstItmzOyPEv: Exit Function
Next
End Function

Function AvzOP(Oy, P As PrpPth) As Variant()
AvzOP = IntozOP(EmpAv, Oy, P)
End Function

Function IntozOP(Into, Oy, P As PrpPth)
Dim O: O = Into: Erase O
Dim I
For Each I In Itr(Oy)
    Push O, Prp(CvObj(I), P)
Next
IntozOP = O
End Function

Function IntAyzOyP(Oy, P As PrpPth) As Integer()
IntAyzOyP = IntozOP(EmpIntAy, Oy, P)
End Function

Function SyzOyPrp(Oy, P As PrpPth) As String()
SyzOyPrp = IntozOP(EmpSy, Oy, P)
End Function

Function OyeFstNEle(Oy, N&)
Dim O: O = Oy
ReDim O(N - 1)
Dim J&
For J = 0 To UB(Oy) - N
    Set O(J) = Oy(N + J)
Next
OyeFstNEle = O
End Function

Function OyeNothing(Oy)
OyeNothing = ResiU(Oy)
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
    If HitNm(Obj.Name, B) Then PushObj OywNm, Obj
Next
End Function

Function OywPredXPTrue(Oy, Xp$, P$)
Dim O, Obj As Object
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If Run(Xp, Obj, P) Then
        PushObj O, Obj
    End If
Next
OywPredXPTrue = O
End Function

Function OywPEv(Oy, P As PrpPth, Ev)
Dim O
   O = Oy
   Erase O
   Dim Obj As Object
   For Each Obj In Itr(Oy)
       If Prp(Obj, P) = Ev Then PushObj O, Obj
   Next
OywPEv = O
End Function
Function OyzItr(Itr) As Variant()
Dim O
For Each O In Itr
    PushObj OyzItr, O
Next
End Function
Function OywPInAy(Oy, P As PrpPth, InAy)
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
For Each I In SyzSS(PP)
    PushI LyzObjPP, I & " " & Prp(Obj, PrpPth(I))
Next
End Function

Private Sub Z_OyDrs()
'ShwWs DrsNewWs(OyDrs(CurrentDb.TableDefs("Z_UpdSeqFld").Fields, "Name Type OrdinalPosition"))
End Sub

Private Sub Z_OyP_Ay()
Dim CdPanAy() As CodePane
Stop
'CdPanAy = Oy(CPj.MdAy).PrpVy("CodePane", CdPanAy)
Stop
End Sub
