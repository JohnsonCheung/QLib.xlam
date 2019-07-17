Attribute VB_Name = "QVb_Dta_Obj"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Obj."
Const DoczP$ = "PrpPth."
Const DoczPn$ = "PrpNm."
Enum EmThw
    EiThwEr
    EiNoThw
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

Function LngAyzOyPrp(Oy, PrpPth$) As Long()
LngAyzOyPrp = CvLngAy(IntozOyPrp(EmpLngAy, Oy, PrpPth))
End Function

Function IntozOyPrp(OInto, Oy, PrpPth$)
Dim O: O = ResiU(OInto)
Dim Obj: For Each Obj In Itr(Oy)
    Push O, Prp(Obj, PrpPth)
Next
IntozOyPrp = O
End Function

Function ObjAddAy(Obj As Object, Oy)
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

Function DrzPrpPthAy(Obj, PthPthAy$()) As Variant()
Const CSub$ = CMod & "DrzObjPrpNy"
If IsNothing(Obj) Then Inf CSub, "Given object is nothing", "PthPthAy", PthPthAy: Exit Function
Dim P: For Each P In PthPthAy
    Dim PrpPth$: PrpPth = P
    Push DrzPrpPthAy, Prp(Obj, PrpPth)
Next
End Function

Function DiczObjPrpPthSy(Obj As Object, PrpPthSy$()) As Dictionary
Dim P, O As New Dictionary
For Each P In PrpPthSy
    O.Add P, Prp(Obj, P)
Next
Set DiczObjPrpPthSy = O
End Function

Function PrpzP(Obj, Prp)
Asg CallByName(Obj, Prp, VbGet), PrpzP
End Function
Function Prp(Obj, PrpPth, Optional ThwEr As EmThw)
Const CSub$ = CMod & "Prp"
'ThwIf_Nothing Obj, CSub
On Error GoTo X
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim PrpSeg$(): PrpSeg = Split(PrpPth, ".")
Dim O
    Set O = Obj
    Dim U%: U = UB(PrpSeg)
    Dim J%: For J = 0 To U - 1     ' U-1 is to skip the last Pth-Seg
        Set O = PrpzP(O, PrpSeg(J)) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next
Asg PrpzP(O, PrpSeg(U)), Prp ' Last Prp may be non-object, so must use 'Asg'
Exit Function
X:
Dim E$: E = Err.Description
If ThwEr = EiThwEr Then
    Thw CSub, "Err", "Er ObjTy PrpPth", E, TypeName(Obj), PrpPth
End If
End Function

