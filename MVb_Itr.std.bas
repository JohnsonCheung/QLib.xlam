Attribute VB_Name = "MVb_Itr"
Option Explicit
Function AvzItr(Itr) As Variant()
AvzItr = IntozItr(Array(), Itr)
End Function

Function ItrAddSfx(Itr, Sfx$) As String()
Dim X
For Each X In Itr
    Push ItrAddSfx, X & Sfx
Next
End Function

Function ItrAddPfx(A, Pfx$) As String()
Dim X
For Each X In A
    Push ItrAddPfx, Pfx & X
Next
End Function

Function ItrClnAy(A)
If A.Count = 0 Then Exit Function
Dim X
For Each X In A
    ItrClnAy = Array(X)
    Exit Function
Next
End Function

Function NItrPrpTrue(A, BoolPrpNm)
Dim O&, X
For Each X In A
    If CallByName(X, BoolPrpNm, VbGet) Then
        O = O + 1
    End If
Next
NItrPrpTrue = O
End Function

Sub DoItrFun(A, DoFun$)
Dim I
For Each I In A
    Run DoFun, I
Next
End Sub

Sub DoItrFunPX(Itr, PX$, P)
Dim X
For Each X In Itr
    Run PX, P, X
Next
End Sub

Sub DoItrFunXP(A, XP$, P)
Dim X
For Each X In A
    Run XP, X, P
Next
End Sub

Function FstItr(A)
Dim X
For Each X In A
    Asg X, FstItr
    Exit Function
Next
End Function

Function FstItrPredXP(A, XP$, P)
Dim X
For Each X In A
    If Run(XP, X, P) Then Asg X, FstItrPredXP: Exit Function
Next
End Function

Function FstItrNm(Itr, Nm) ' Return first element in Itr-A with name eq Nm
Dim O
For Each O In Itr
    If ObjNm(O) = Nm Then
        Set FstItrNm = O
        Exit Function
    End If
Next
Set FstItrNm = Nothing
End Function
Function FstItrPEv(Itr, P, Ev) 'Return first element in Itr-A with its Prp-P eq to V
Dim O
For Each O In Itr
    If ObjPrp(O, P) = Ev Then Set FstItrPEv = O: Exit Function
Next
Set FstItrPEv = Nothing
End Function

Function FstItrTrueP(Itr, TruePrp) 'Return first element in Itr with its Prp-P being true
Set FstItrTrueP = FstItrPEv(Itr, TruePrp, True)
End Function
Function HasItrTrueP(Itr, TruePrp) As Boolean
Dim Obj
For Each Obj In Itr
    If ObjPrp(Obj, TruePrp) Then HasItrTrueP = True: Exit Function
Next
End Function
Function HasItn(Itr, Nm) As Boolean
Dim O
For Each O In Itr
    If O.Name = Nm Then HasItn = True: Exit Function
Next
End Function

Function HasItrPEv(A, P, Ev) As Boolean
Dim X
For Each X In A
    If ObjPrp(X, P) = Ev Then HasItrPEv = True: Exit Function
Next
End Function

Function HasItrTruePrp(A, P) As Boolean
Dim X
For Each X In A
    If ObjPrp(X, P) Then HasItrTruePrp = True: Exit Function
Next
End Function

Function IsEqNmItr(A, B)
IsEqNmItr = IsSamAy(Itn(A), Itn(B))
End Function

Function ItrMap(Itr, Map$) As Variant()
ItrMap = IntozItrMap(EmpAv, Itr, Map)
End Function

Function IntozItrMap(OInto, Itr, Map$)
Dim O: O = OInto
Erase O
Dim X
For Each X In Itr
    Push O, Run(Map, X)
Next
IntozItrMap = O
End Function

Function SyzItrMap(Itr, Map$) As String()
SyzItrMap = IntozItrMap(EmpSy, Itr, Map)
End Function

Function MaxItrPrp(A, P)
Dim X, O
For Each X In A
    O = Max(O, ObjPrp(X, P))
Next
MaxItrPrp = O
End Function
Function NyOy(A) As String()
Dim I
For Each I In Itr(A)
    PushI NyOy, ObjNm(I)
Next
End Function
Function ItnPEv(Itr, WhPrp, Ev) As String()
Dim Obj
For Each Obj In Itr
    If ObjPrp(Obj, WhPrp) = Ev Then PushI ItnPEv, ObjNm(Obj)
Next
End Function
Function SyPrp(Itr, P) As String()
Dim Obj
For Each Obj In Itr
    PushI SyPrp, ObjPrp(Obj, P)
Next
End Function
Function VyzItr(Itr) As Variant()
Dim Obj
For Each Obj In Itr
    PushI VyzItr, Obj.Value
Next
End Function
Function Itn(Itr) As String()
Dim I
For Each I In Itr
    PushI Itn, ObjNm(I)
Next
End Function

Function IsAllFalse_ItrPred(Itr, Pred$) As Boolean
Dim I
For Each I In Itr
    If Run(Pred, I) Then Exit Function
Next
IsAllFalse_ItrPred = True
End Function

Function IsAllTrue_ItrPred(A, Pred$) As Boolean
Dim I
For Each I In A
    If Not Run(Pred, I) Then Exit Function
Next
IsAllTrue_ItrPred = True
End Function

Function IsSomFalse_ItrPred(Itr, Pred$) As Boolean
Dim I
For Each I In Itr
    If Not Run(Pred, I) Then IsSomFalse_ItrPred = True: Exit Function
Next
End Function

Function IsSomTrue_ItrPred(Itr, Pred$) As Boolean
Dim I
For Each I In Itr
    If Run(Pred, I) Then IsSomTrue_ItrPred = True: Exit Function
Next
End Function
Function SyItrPrp(Itr, P) As String()
Stop
SyItrPrp = IntozItrPrp(Itr, P, EmpSy)
End Function
Function AvItrPrp(Itr, P) As Variant()
AvItrPrp = IntozItrPrp(Itr, P, EmpAv)
End Function
Function IntozItrPrpTrue(Into, Itr, P)
IntozItrPrpTrue = AyCln(Into)
Dim I
For Each I In Itr
    If ObjPrp(I, P) Then
        Push IntozItrPrpTrue, I
    End If
Next
End Function
Function IntozItrPEv(Into, Itr, P, Ev)
IntozItrPEv = AyCln(Into)
Dim Obj
For Each Obj In Itr
    If ObjPrp(Obj, P) = Ev Then PushObj IntozItrPEv, Obj
Next
End Function
Function IntozItrPrp(Into, Itr, P)
IntozItrPrp = AyCln(Into)
Dim I
For Each I In Itr
    Push IntozItrPrp, ObjPrp(I, P)
Next
End Function

Function AvItrValue(A) As Variant()
AvItrValue = AvItrPrp(A, "Value")
End Function

Function ItrPrp_WhTrue_Into(A, P, Into)
Dim O: O = Into: Erase O
Dim Obj
For Each Obj In A
    If CallByName(Obj, P, VbGet) Then
        PushObj Into, Obj
    End If
Next
ItrPrp_WhTrue_Into = O
End Function

Function ItrwPrpEqval(A, Prp, EqVal)
ItrwPrpEqval = ItrClnAy(A)
Dim O
For Each O In A
    If ObjPrp(O, Prp) = EqVal Then PushObj ItrwPrpEqval, O
Next
ItrwPrpEqval = O
End Function

Function ItrwPrpTrue(A, P)
ItrwPrpTrue = ItrClnAy(A)
Dim O
For Each O In A
    If ObjPrp(O, P) Then
        Push ItrwPrpTrue, O
    End If
Next
End Function

Function ItrwNm(A, B As WhNm)
ItrwNm = ItrClnAy(A)
Dim O
For Each O In A
    If HitNm(ObjNm(O), B) Then
        Push ItrwNm, O
    End If
Next
End Function

Private Sub ZZ()
Dim A As Variant
Dim B$
Dim C As RegExp
Dim D$()
Dim E As WhNm
AvzItr A
ItrClnAy A
DoItrFun A, B
DoItrFun A, B
DoItrFunPX A, B, A
DoItrFunXP A, B, A
FstItr A
FstItr A
SyzItrMap A, B
MaxItrPrp A, A
Itn A
IsAllFalse_ItrPred A, B
IsAllTrue_ItrPred A, B
IsSomFalse_ItrPred A, B
IsSomTrue_ItrPred A, B
ItrwPrpTrue A, A
End Sub

Private Sub Z()
End Sub
Function NItrPEv&(Itr, P, Ev)
Dim O&, Obj
For Each Obj In Itr
    If ObjPrp(Obj, P) = Ev Then O = O + 1
Next
NItrPEv = O
End Function
