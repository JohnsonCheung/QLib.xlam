Attribute VB_Name = "QVb_Itr"
Option Explicit
Private Const CMod$ = "MVb_Itr."
Private Const Asm$ = "QVb"
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

Function ItrAddPfx(Itr, Pfx$) As String()
Dim X
For Each X In Itr
    Push ItrAddPfx, Pfx & X
Next
End Function

Function ItrClnAy(Itr)
If Itr.Count = 0 Then Exit Function
Dim X
For Each X In Itr
    ItrClnAy = Array(X)
    Exit Function
Next
End Function

Function NItrPrpTrue(Itr, BoolPrpNm)
Dim O&, X
For Each X In Itr
    If CallByName(X, BoolPrpNm, VbGet) Then
        O = O + 1
    End If
Next
NItrPrpTrue = O
End Function

Sub DoItrFun(Itr, DoFun$)
Dim I
For Each I In Itr
    Run DoFun, I
Next
End Sub

Sub DoItrFunPX(Itr, PX$, P)
Dim X
For Each X In Itr
    Run PX, P, X
Next
End Sub

Sub DoItrFunXP(Itr, XP$, P)
Dim X
For Each X In Itr
    Run XP, X, P
Next
End Sub

Function FstItm(Itr)
Dim X
For Each X In Itr
    Asg X, FstItm
    Exit Function
Next
End Function

Function FstItmPredXP(Ay, XP$, P$)
Dim X
For Each X In Ay
    If Run(XP, X, P) Then Asg X, FstItmPredXP: Exit Function
Next
End Function

Function FstItmzNm(Itr, Nm$) ' Return first element in Itr-Itr with name eq Nm
Dim O
For Each O In Itr
    If ObjNm(O) = Nm Then
        Set FstItmzNm = O
        Exit Function
    End If
Next
Set FstItmzNm = Nothing
End Function
Function FstItmPEv(Itr, P As PrpPth, Ev) 'Return first element in Itr-Itr with its Prp-P eq to V
Dim I, Obj As Object
For Each I In Itr
    If Prp(CvObj(I), P) = Ev Then Set FstItmPEv = I: Exit Function
Next
Set FstItmPEv = Nothing
End Function
Function FstItn(Itr, Nm$) 'Return first element in Itr with its PrpNm=Nm being true
Set FstItn = FstItmPEv(Itr, PrpPth("Name"), Nm)
End Function

Function FstItmTrueP(Itr, TruePrpPth As PrpPth) 'Return first element in Itr with its Prp-P being true
Set FstItmTrueP = FstItmPEv(Itr, TruePrpPth, True)
End Function

Function HasItn(Itr, Nm) As Boolean
Dim O
For Each O In Itr
    If O.Name = Nm Then HasItn = True: Exit Function
Next
End Function

Function HasItrPEv(Itr, P$, Ev) As Boolean
Dim I
For Each I In Itr
    If Prp(CvObj(I), PrpPth(P)) = Ev Then HasItrPEv = True: Exit Function
Next
End Function

Function HasItrTruePrp(Itr, P$) As Boolean
Dim I
For Each I In Itr
    If Prp(CvObj(I), PrpPth(P)) Then HasItrTruePrp = True: Exit Function
Next
End Function

Function IsEqNmItr(Itr, B)
IsEqNmItr = IsSamAy(Itn(Itr), Itn(B))
End Function

Function AvzItrMap(Itr, Map$) As Variant()
AvzItrMap = IntozItrMap(EmpAv, Itr, Map)
End Function

Function IntozMapAy(OInto, Ay, Map$)
IntozMapAy = IntozItrMap(OInto, Itr(Ay), Map)
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

Private Sub Z_PrpVy()
Vc PrpVy(CPj.VBComponents, PrpPth("CodeModule.CountOfLines"))
End Sub

Function PrpVy(Itr, P As PrpPth) As Variant()
Dim O As Object
For Each O In Itr
    Push PrpVy, Prp(O, P)
Next
End Function
Function MaxzItrPrp(Itr, P$)
Dim I, O
For Each I In Itr
    O = Max(O, Prp(CvObj(I), PrpPth(P)))
Next
MaxzItrPrp = O
End Function

Function NyzItr(Itr) As String()
NyzItr = Itn(Itr)
End Function
Function ItnPEv(Itr, WhPrp$, Ev) As String()
Dim Obj As Object
For Each Obj In Itr
    If Prp(Obj, PrpPth(WhPrp)) = Ev Then PushI ItnPEv, ObjNm(Obj)
Next
End Function
Function NyzOy(Oy) As String()
NyzOy = Itn(Itr(Oy))
End Function
Function PvyzItr(Itr, P$) As Variant()
Dim Obj As Object
For Each Obj In Itr
    Push PvyzItr, Prp(Obj, PrpPth(P))
Next
End Function
Function Itn(Itr) As String()
Dim I
For Each I In Itr
    PushI Itn, ObjNm(I)
Next
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

Function PrpSy(Itr, P$) As String()
PrpSy = SyzItrPrp(Itr, P)
End Function

Function SyzItrPrp(Itr, P$) As String()
SyzItrPrp = IntozItrPrp(EmpSy, Itr, P)
End Function

Function AvzItrPrp(Itr, P$) As Variant()
AvzItrPrp = IntozItrPrp(EmpAv, Itr, P)
End Function

Function IntozItrwPredTrue(Into, Itr, P As IPred)
IntozItrwPredTrue = Resi(Into)
Dim Obj As Object
For Each Obj In Itr
    If P.Pred(Obj) Then
        Push IntozItrwPredTrue, Obj
    End If
Next
End Function

Function IntozItrwPEv(Into, Itr, P$, Ev)
IntozItrwPEv = Resi(Into)
Dim Obj As Object
For Each Obj In Itr
    If Prp(Obj, PrpPth(P)) = Ev Then PushObj IntozItrwPEv, Obj
Next
End Function
Function Into(OInto, Itr)
Into = Resi(OInto)
Dim I
For Each I In Itr
    Push Into, I
Next
End Function
Function IntozItrPrp(Into, Itr, P$)
IntozItrPrp = Resi(Into)
Dim Obj As Object
For Each Obj In Itr
    Push IntozItrPrp, Prp(Obj, PrpPth(P))
Next
End Function

Function AvItrValue(Itr) As Variant()
AvItrValue = AvzItrPrp(Itr, "Value")
End Function

Function ItrwPrpEqval(Itr, P$, EqVal)
ItrwPrpEqval = ItrClnAy(Itr)
Dim O As Object
For Each O In Itr
    If Prp(O, PrpPth(P)) = EqVal Then PushObj ItrwPrpEqval, O
Next
ItrwPrpEqval = O
End Function

Function ItrwPrpTrue(Itr, P$)
ItrwPrpTrue = ItrClnAy(Itr)
Dim O As Object
For Each O In Itr
    If Prp(O, PrpPth(P)) Then
        Push ItrwPrpTrue, O
    End If
Next
End Function

Function ItrwNm(Itr, B As WhNm)
ItrwNm = ItrClnAy(Itr)
Dim O
For Each O In Itr
    If HitNm(ObjNm(O), B) Then
        Push ItrwNm, O
    End If
Next
End Function

Private Sub ZZ()
Dim Itr As Variant
Dim B$
Dim C As RegExp
Dim D$()
Dim E As WhNm
AvzItr Itr
ItrClnAy Itr
DoItrFun Itr, B
DoItrFun Itr, B
DoItrFunPX Itr, B, Itr
DoItrFunXP Itr, B, Itr
FstItm Itr
FstItm Itr
Itn Itr
End Sub

Function NItrPEv&(Itr, P As PrpPth, Ev)
Dim O&, Obj As Object
For Each Obj In Itr
    If Prp(Obj, P) = Ev Then O = O + 1
Next
NItrPEv = O
End Function
Function VyzItrp(Itr, P As PrpPth) As String()
Dim Obj As Object
For Each Obj In Itr
    Push VyzItrp, Prp(Obj, P)
Next
End Function


Function PrpNy(Itr) As String()
PrpNy = Itn(Itr.Properties)
End Function

