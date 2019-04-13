Attribute VB_Name = "MVb_Obj"
Option Explicit
Const CMod$ = "MVb__Obj."
Enum eThwOpt
    eeThw
    eeNoThwInf
    eeNoThwNoInf
End Enum
Function IsEqObj(A, B) As Boolean
IsEqObj = ObjPtr(A) = ObjPtr(B)
End Function

Function IsEqVar(A, B) As Boolean
IsEqVar = VarPtr(A) = VarPtr(B)
End Function

Function IntozOy(OInto, Oy)
Erase OInto
Dim O, I
For Each I In Itr(Oy)
    PushObj OInto, I
Next
End Function

Function LngAyzOyPrp(Oy, Prp) As Long()
LngAyzOyPrp = CvLngAy(IntozOyPrp(EmpLngAy, Oy, Prp))
End Function

Function IntozOyPrp(OInto, Oy, Prp)
Dim O, I
O = AyCln(OInto)
For Each I In Itr(Oy)
    Push O, Prp(I, Prp)
Next
IntozOyPrp = O
End Function

Function ObjAddAy(Obj, Oy)
Dim O: O = Oy
Erase O
PushObj O, Obj
PushObjAy O, Oy
ObjAddAy = O
End Function

Function ObjNm$(A)
If IsNothing(A) Then ObjNm = "#Obj Is Nothing#": Exit Function
On Error GoTo X
ObjNm = A.Name
Exit Function
X:
ObjNm = "#" & Err.Description & "#"
End Function

Function DrzObj(Obj, PrpPthAy$()) As Variant()
Const CSub$ = CMod & "DrzObjPrpNy"
If IsNothing(Obj) Then Inf CSub, "Given object is nothing", "PrpPthAy", PrpPthAy: Exit Function
Dim PrpPth
For Each PrpPth In PrpPthAy
    Push DrzObj, Prp(Obj, PrpPth)
Next
End Function

Function LyzObjPP(Obj, PP) As String()
LyzObjPP = LyzObjPP(Obj, Ny(PP))
End Function

Function ObjStr$(Obj)
On Error GoTo X
ObjStr = Obj.ToStr: Exit Function
X: ObjStr = QuoteSq(TypeName(Obj))
End Function

Private Sub ZZZ_Prp()
Dim Act$: Act = Prp(Excel.Application.Vbe.ActiveVBProject, "FileName Name")
Ass Act = "C:\Users\user\Desktop\Vba-Lib-1\QVb.xlam|QVb"
End Sub

Function PrpzNm(Obj, P) ' P is PrpNm (Nm cannot have Dot
On Error Resume Next
PrpzNm = CallByName(Obj, P, VbGet)
End Function

Function Prp(Obj, PrpPth, Optional Thw As eThwOpt)
Const CSub$ = CMod & "Prp"
'ThwNothing Obj, CSub
On Error GoTo X
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim P$()
    P = Split(PrpPth, ".")
Dim O
    Dim J%, U%
    Set O = Obj
    U = UB(P)
    For J = 0 To U - 1      ' U-1 is to skip the last Pth-Seg
        Set O = CallByName(O, P(J), VbGet) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next
Asg CallByName(O, P(U), VbGet), Prp ' Last Prp may be non-object, so must use 'Asg'
Exit Function
X:
Dim E$: E = Err.Description
ThwOpt Thw, CSub, "Err", "Er ObjTy PrpPth", E, TypeName(Obj), PrpPth
End Function


