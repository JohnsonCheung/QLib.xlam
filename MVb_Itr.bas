Attribute VB_Name = "MVb_Itr"
Option Explicit
Function ObjVyzItr(Itr) As Variant()
Dim Obj
For Each Obj In Itr
    PushI ObjVyzItr, Obj.Value
Next
End Function

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
    If Prp(O, P) = Ev Then Set FstItrPEv = O: Exit Function
Next
Set FstItrPEv = Nothing
End Function
Function FstItn(Itr, Nm) 'Return first element in Itr with its PrpNm=Nm being true
Set FstItn = FstItrPEv(Itr, "Name", Nm)
End Function

Function FstItrTrueP(Itr, TruePrp) 'Return first element in Itr with its Prp-P being true
Set FstItrTrueP = FstItrPEv(Itr, TruePrp, True)
End Function
Function HasItrTrueP(Itr, TruePrp) As Boolean
Dim Obj
For Each Obj In Itr
    If Prp(Obj, TruePrp) Then HasItrTrueP = True: Exit Function
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
    If Prp(X, P) = Ev Then HasItrPEv = True: Exit Function
Next
End Function

Function HasItrTruePrp(A, P) As Boolean
Dim X
For Each X In A
    If Prp(X, P) Then HasItrTruePrp = True: Exit Function
Next
End Function

Function IsEqNmItr(A, B)
IsEqNmItr = IsSamAy(Itn(A), Itn(B))
End Function

Function AvzItrMap(Itr, Map$) As Variant()
AvzItrMap = IntozAvzItrMap(EmpAv, Itr, Map)
End Function

Function IntozAyMap(OInto, Ay, Map$)
IntozAyMap = IntozAvzItrMap(OInto, Itr(Ay), Map)
End Function

Function IntozAvzItrMap(OInto, Itr, Map$)
Dim O: O = OInto
Erase O
Dim X
For Each X In Itr
    Push O, Run(Map, X)
Next
IntozAvzItrMap = O
End Function

Function SyzAvzItrMap(Itr, Map$) As String()
SyzAvzItrMap = IntozAvzItrMap(EmpSy, Itr, Map)
End Function

Function MaxItrPrp(A, P)
Dim X, O
For Each X In A
    O = Max(O, Prp(X, P))
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
    If Prp(Obj, WhPrp) = Ev Then PushI ItnPEv, ObjNm(Obj)
Next
End Function
Function SyPrp(Itr, P) As String()
Dim Obj
For Each Obj In Itr
    PushI SyPrp, Prp(Obj, P)
Next
End Function
Function NyzOy(Oy) As String()
NyzOy = Itn(Itr(Oy))
End Function
Function PrpVyzItr(Itr, PrpNm) As Variant()

End Function
Function Itn(Itr) As String()
Dim I
For Each I In Itr
    PushI Itn, ObjNm(I)
Next
End Function

Function IsAllFalsezItrPred(Itr, Pred$) As Boolean
Dim I
For Each I In Itr
    If Run(Pred, I) Then Exit Function
Next
IsAllFalsezItrPred = True
End Function

Function IsAllTruezItrPred(Itr, Pred$) As Boolean
Dim I
For Each I In Itr
    If Not Run(Pred, I) Then Exit Function
Next
IsAllTruezItrPred = True
End Function

Function IsSomFalsezItrPred(Itr, Pred$) As Boolean
Dim I
For Each I In Itr
    If Not Run(Pred, I) Then IsSomFalsezItrPred = True: Exit Function
Next
End Function

Function IsSomTruezItrPred(Itr, Pred$) As Boolean
Dim I
For Each I In Itr
    If Run(Pred, I) Then IsSomTruezItrPred = True: Exit Function
Next
End Function
Function SyzItrPrp(Itr, P) As String()
SyzItrPrp = IntozItrPrp(EmpSy, Itr, P)
End Function

Function AvzItrPrp(Itr, P) As Variant()
AvzItrPrp = IntozItrPrp(Itr, P, EmpAv)
End Function
Function IntozItrPrpTrue(Into, Itr, P)
IntozItrPrpTrue = AyCln(Into)
Dim I
For Each I In Itr
    If Prp(I, P) Then
        Push IntozItrPrpTrue, I
    End If
Next
End Function

Function IntozItrPEv(Into, Itr, P, Ev)
IntozItrPEv = AyCln(Into)
Dim Obj
For Each Obj In Itr
    If Prp(Obj, P) = Ev Then PushObj IntozItrPEv, Obj
Next
End Function
Function IntozItrPrp(Into, Itr, P)
IntozItrPrp = AyCln(Into)
Dim I
For Each I In Itr
    Push IntozItrPrp, Prp(I, P)
Next
End Function

Function AvItrValue(A) As Variant()
AvItrValue = AvzItrPrp(A, "Value")
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
    If Prp(O, Prp) = EqVal Then PushObj ItrwPrpEqval, O
Next
ItrwPrpEqval = O
End Function

Function ItrwPrpTrue(A, P)
ItrwPrpTrue = ItrClnAy(A)
Dim O
For Each O In A
    If Prp(O, P) Then
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
SyzAvzItrMap A, B
MaxItrPrp A, A
Itn A
IsAllFalsezItrPred A, B
IsAllTruezItrPred A, B
IsSomFalsezItrPred A, B
ItrwPrpTrue A, A
End Sub

Private Sub Z()
End Sub
Function NItrPEv&(Itr, P, Ev)
Dim O&, Obj
For Each Obj In Itr
    If Prp(Obj, P) = Ev Then O = O + 1
Next
NItrPEv = O
End Function
Function VyzItrPrp(Itr, PrpPth) As String()
Dim Obj
For Each Obj In Itr
    Push VyzItrPrp, Prp(Obj, PrpPth)
Next
End Function


Function PrpNy(A) As String()
PrpNy = Itn(A.Properties)
End Function

