Attribute VB_Name = "MVb_Obj"
Option Explicit
Const CMod$ = "MVb__Obj."
Enum eThwOpt
    eThw
    eNoThwInfo
    eNoThwNoInfo
End Enum
Function IsEqObj(A, B) As Boolean
IsEqObj = ObjPtr(A) = ObjPtr(B)
End Function
Function IntozOy(OInto, Oy)
Erase OInto
Dim O, I
For Each I In Itr(Oy)
    PushObj OInto, I
Next
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

Function DrzObjPrpNy(Obj, PrpNy$()) As Variant()
Const CSub$ = CMod & "DrzObjPrpNy"
If IsNothing(Obj) Then Info CSub, "Given object is nothing", "PrpNy", PrpNy: Exit Function
Dim I
For Each I In PrpNy
    Push DrzObjPrpNy, ObjPrp(Obj, I)
Next
End Function
Function LyzObjPrpNy(Obj, B$()) As String()
LyzObjPrpNy = LyzNyAv(B, DrzObjPrpNy(Obj, B))
End Function
Function LyzObjPP(Obj, PP) As String()
LyzObjPP = LyzObjPrpNy(Obj, Ny(PP))
End Function
Function DrzObjPP(Obj, PP$) As Variant()
DrzObjPP = DrzObjPrpNy(Obj, Ny(PP))
End Function

Function ObjPrp(A, PrpPth, Optional Thw As eThwOpt)
Const CSub$ = CMod & "ObjPrp"
'ThwNothing A, CSub
On Error GoTo X
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim P$()
    P = Split(PrpPth, ".")
Dim O
    Dim J%, U%
    Set O = A
    U = UB(P)
    For J = 0 To U - 1      ' U-1 is to skip the last Pth-Seg
        Set O = CallByName(O, P(J), VbGet) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next
Asg CallByName(O, P(U), VbGet), ObjPrp ' Last Prp may be non-object, so must use 'Asg'
Exit Function
X:
Dim E$: E = Err.Description
ThwOpt Thw, CSub, "Err", "Er ObjTy PrpPth", E, TypeName(A), PrpPth
End Function

Function Obj_ToStr$(A)
'On Error GoTo X
Obj_ToStr = A.ToStr: Exit Function
'X: Obj_ToStr = QuoteSq(TypeName(A))
End Function

Private Sub ZZZ_ObjPrp()
Dim Act$: Act = ObjPrp(Excel.Application.Vbe.ActiveVBProject, "FileName Name")
Ass Act = "C:\Users\user\Desktop\Vba-Lib-1\QVb.xlam|QVb"
End Sub


Function Prp(Obj, P)
On Error Resume Next
Prp = CallByName(Obj, P, VbGet)
End Function

