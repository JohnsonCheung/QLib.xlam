Attribute VB_Name = "MxObj"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxObj."
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

