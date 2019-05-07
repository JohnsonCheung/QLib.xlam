Attribute VB_Name = "QVb_Obj"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Obj."
Const DoczP$ = "PrpPth."
Const DoczPn$ = "PrpNm."
Type Nm: Nm As String: End Type
Type PrpPth: P As String: End Type
Enum EmThw
    EiThw
    EiNoThwInf
    EiNoThwQuiet
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
O = Resi(OInto)
For Each I In Itr(Oy)
    Push O, Prp(I, Prp)
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

Function DrzObjPrpPthSy(Obj As Object, PrpPthSy$()) As Variant()
Const CSub$ = CMod & "DrzObjPrpNy"
If IsNothing(Obj) Then Inf CSub, "Given object is nothing", "PrpPthSy", PrpPthSy: Exit Function
Dim PrpPth
For Each PrpPth In PrpPthSy
    Push DrzObjPrpPthSy, Prp(Obj, CStr(PrpPth))
Next
End Function

Function DiczObjPP(Obj As Object, PP$) As Dictionary
Set DiczObjPP = DiczObjPrpPthSy(Obj, Ny(PP))
End Function
Function DiczObjPrpPthSy(Obj As Object, PrpPthNy$()) As Dictionary
Dim PrpPth, O As New Dictionary
For Each PrpPth In PrpPthNy
    O.Add PrpPth, Prp(Obj, CStr(PrpPth))
Next
Set DiczObjPrpPthSy = O
End Function
Function ObjToStr$(Obj As Excel.Application)
On Error GoTo X
ObjToStr = Obj.ToStr: Exit Function
X: ObjToStr = QuoteSq(TypeName(Obj))
End Function

Private Sub ZZZ_Prp()
Dim Act$: Act = Prp(Excel.Application.Vbe.ActiveVBProject, "FileName Name")
Ass Act = "C:\Users\user\Desktop\Vba-Lib-1\QVb.xlam|QVb"
End Sub
Function PrpzNm(Obj As Object, N As Nm, Optional Thw As EmThw)

End Function
Function Prp(Obj As Object, P As PrpPth, Optional Thw As EmThw)
Const CSub$ = CMod & "Prp"
'ThwIfNothing Obj, CSub
On Error GoTo X
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim SegSy$()
    SegSy = Split(P.P, ".")
Dim O
    Dim J%, U%
    Set O = Obj
    U = UB(SegSy)
    For J = 0 To U - 1      ' U-1 is to skip the last Pth-Seg
        Set O = CallByName(O, SegSy(J), VbGet) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next
Asg CallByName(O, SegSy(U), VbGet), Prp ' Last Prp may be non-object, so must use 'Asg'
Exit Function
X:
Dim E$: E = Err.Description
ThwOpt Thw, CSub, "Err", "Er ObjTy PrpPth", E, TypeName(Obj), P.P
End Function

