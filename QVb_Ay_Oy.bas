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

Function FstzOyEq(Oy, PrpPth, V)
Set FstzOyEq = FstzItrEq(Itr(Oy), PrpPth, V)
End Function

Function AvzOyP(Oy, PrpPth) As Variant()
AvzOyP = IntozOyP(EmpAv, Oy, PrpPth)
End Function

Function IntozOyP(Into, Oy, PrpPth)
Dim O: O = Into: Erase O
Dim Obj: For Each Obj In Itr(Oy)
    Push O, Prp(Obj, PrpPth)
Next
IntozOyP = O
End Function

Function IntAyzOyP(Oy, PrpPth) As Integer()
IntAyzOyP = IntozOyP(EmpIntAy, Oy, PrpPth)
End Function

Function SyzOyP(Oy, PrpPth) As String()
SyzOyP = IntozOyP(EmpSy, Oy, PrpPth)
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

Function FstzObj(Oy, PrpPth$, V)
'Ret : Fst Obj in @Oy having @PrpPth = @V
Dim Obj: For Each Obj In Itr(Oy)
    If Prp(Obj, PrpPth) = V Then Asg Obj, FstzObj: Exit Function
Next
End Function
Function OyzItr(Itr) As Variant()
Dim O
For Each O In Itr
    PushObj OyzItr, O
Next
End Function
Function OywIn(Oy, PrpPth, InAy)
Dim Obj As Object, O
If Si(Oy) = 0 Or Si(InAy) Then OywIn = Oy: Exit Function
O = Oy
Erase O
For Each Obj In Itr(Oy)
    If HasEle(InAy, Prp(Obj, PrpPth)) Then PushObj O, Obj
Next
OywIn = O
End Function

Function LyzObjPP(Obj As Object, PP$) As String()
Dim PrpPth: For Each PrpPth In SyzSS(PP)
    PushI LyzObjPP, PrpPth & " " & Prp(Obj, PrpPth)
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

'
