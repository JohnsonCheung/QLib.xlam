Attribute VB_Name = "QVb_Thw"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Thw."
Type CfgInf
    ShwInf As Boolean
    ShwTim As Boolean
End Type
Type Cfg
    Inf As CfgInf
End Type
Public Cfg As Cfg
Sub ThwIfNegEle(Ay, Fun$)
Const CSub$ = CMod & "ThwIfNEgEle"
Dim I, J&, O$()
For Each I In Itr(Ay)
    If I < 0 Then
        PushI O, J & ": " & I
        J = J + 1
    End If
Next
If Si(O) > 0 Then
    Thw CSub, "In [Ay], there are [negative-element (Ix Ele)]", Ay, O
End If
End Sub

Sub ThwIfNE(A, B, Optional ANm$ = "A", Optional BNm$ = "B")
Const CSub$ = CMod & "ThwIfNE"
ThwIfDifTy A, B, ANm, BNm
Dim IsBothLines As Boolean
    IsBothLines = IsLines(A) Or IsLines(B)
Select Case True
Case IsBothLines: If A <> B Then CmpLines CStr(A), CStr(B), Hdr:=FmtQQ("Lines ? ? not eq.", ANm, BNm): Stop: Exit Sub
Case IsStr(A):    If A <> B Then CmpStr CStr(A), CStr(B), Hdr:=FmtQQ("String ? ? not eq.", ANm, BNm): Stop: Exit Sub
Case IsDic(A):    If Not IsEqDic(CvDic(A), CvDic(B)) Then BrwCmpDicAB CvDic(A), CvDic(B): Stop: Exit Sub
Case IsArray(A):  ThwIfNEAy A, B, ANm, BNm
Case IsObject(A): If ObjPtr(A) <> ObjPtr(B) Then Thw CSub, "Two object are diff", FmtQQ("Ty-? Ty-?", ANm, BNm), TypeName(A), TypeName(B)
Case Else:
    If A <> B Then
        Thw CSub, "A B NE", "A B", A, B
        Exit Sub
    End If
End Select
End Sub
Private Sub ThwIfNEAy(AyA, AyB, ANm$, BNm$)
ThwIfDifSi AyA, AyB, ANm, BNm
ThwIfDifTy AyA, AyB, ANm, BNm
Dim J&, A
For Each A In Itr(AyA)
    If Not IsEq(A, AyB(J)) Then
        Dim NN$: NN = FmtQQ("AyTy AySi Dif-At Ay-?-Ele-?-Ty Ay-?-Ele-?-Ty Ay-?-Ele-Val Ay-?-Ele-Val Ay-? Ay-?", ANm, J, BNm, J, ANm, BNm, ANm, BNm)
        Thw CSub, "There is ele in 2 Ay are diff", NN, TypeName(AyA), Si(AyA), J, TypeName(A), TypeName(AyB(J)), A, AyB(J), AyA, AyB
    End If
    J = J + 1
Next
End Sub
Sub ThwIfDifTy(A, B, Optional ANm$ = "A", Optional BNm$ = "B")
If TypeName(A) = TypeName(B) Then Exit Sub
Dim NN$
NN = FmtQQ("?-TyNm ?-TyNm", ANm, BNm)
Thw CSub, "Type Diff", NN, TypeName(A), TypeName(B)
End Sub


Sub ThwIfDifSi(A, B, Fun$)
If Si(A) <> Si(B) Then Thw Fun, "Si-A <> Si-B", "Si-A Si-B", Si(A), Si(B)
End Sub

Sub ThwIfObjNE(A, B, Fun$, Msg$, Nav())
If IsEqObj(A, B) Then ThwNav Fun, Msg, Nav
End Sub

Sub ThwIfNoSrt(Ay, Fun$)
If IsSrtAy(Ay) Then Thw Fun, "Array should be sorted", "Ay-Ty Ay", TypeName(Ay), Ay
End Sub
Sub ThwOpt(Thw As EmThw, Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Select Case Thw
Case EiNoThwInf: InfNav Fun, Msg, Nav
Case EiNoThwQuiet:
Case Else:   ThwNav Fun, Msg, Nav
End Select
End Sub

Sub Thw(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Const PgmErBox$ = _
"*******************" & vbCrLf & _
"** Program error **" & vbCrLf & _
"*******************" & vbCrLf
ThwNav Fun, PgmErBox & Msg, Nav
End Sub

Sub ThwNav(Fun$, Msg$, Nav())
BrwAy LyzFunMsgNav(Fun, Msg, Nav)
Halt
End Sub

Sub Ass(A As Boolean)
Debug.Assert A
End Sub
Sub ThwIfNothing(A, Fun$)
If Not IsNothing(A) Then Exit Sub
Thw Fun, "Given parameter is nothing"
End Sub
Sub ThwIfNotAy(A, Fun$)
If IsArray(A) Then Exit Sub
Thw Fun, "Given parameter should be array, but now TypeName=" & TypeName(A)
End Sub
Sub ThwIfNEver(Fun$, Optional Msg$ = "Program should not reach here")
Thw Fun, Msg
End Sub

Sub Halt(Optional Fun$)
Err.Raise -1, Fun, "Please check messages opened in notepad"
End Sub

Sub Done()
MsgBox "Done"
End Sub

Function AddNNAv(Nav(), NN$, Av()) As Variant()
Dim O(): O = Nav
If Si(O) = 0 Then
    PushI O, NN
Else
    O(0) = O(0) & " " & NN
End If
PushAy O, Av
AddNNAv = O
End Function

Function AddNmV(Nav(), Nm$, V) As Variant()
AddNmV = AddNNAv(Nav, Nm, Av(V))
End Function
Sub ThwIfErMsg(Er$(), Fun$, Msg$, ParamArray Nap())
If Si(Er) = 0 Then Exit Sub
Dim Nav(): Nav = Nap
ThwNav Fun, Msg, AddNmV(Nav, "Er", Er)
End Sub
Sub ThwIfEr(Er$(), Fun$)
If Si(Er) = 0 Then Exit Sub
BrwAy AddSy(LyzFunMsgNap(Fun, ""), Er)
Halt
End Sub
Sub ThwLoopingTooMuch(Fun$)
Thw Fun, "Looping too much"
End Sub
Sub ThwPmEr(ValzPm, Fun$, Optional MsgWhyPmEr$ = "Invalid value")
Thw Fun, "Parameter error: " & MsgWhyPmEr, "Pm-Type Pm-Val", TypeName(ValzPm), FmtV(ValzPm)
End Sub

Sub D(Optional A)
Select Case True
Case IsMissing(A): Debug.Print
Case IsArray(A): DmpAy A
Case IsDic(A):   DmpDic CvDic(A), True
Case IsRel(A):   CvRel(A).Dmp
Case IsAset(A):  CvAset(A).Dmp
Case Else: Debug.Print A
End Select
End Sub
Sub Dmp(A)
D A
End Sub
Sub DmpTy(A)
Debug.Print TypeName(A)
End Sub

Sub DmpAyWithIx(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print J; ": "; Ay(J)
Next
End Sub
Sub DmpAy(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print Ay(J)
Next
End Sub

Sub InfLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
D LinzFunMsgNav(Fun, Msg, Nav)
End Sub
Sub InfNav(Fun$, Msg$, Nav())
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub
Sub Inf(Fun$, Msg$, ParamArray Nap())
If Not Cfg.Inf.ShwInf Then Exit Sub
Dim Nav(): Nav = Nap
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub WarnLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Debug.Print LinzFunMsgNav(Fun, Msg, Nav)
End Sub

Sub Warn(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub

Private Sub Z_LyzObjPP()
Dim Obj As Object, PP$
GoSub T0
Exit Sub
T0:
    Set Obj = New DAO.Field
    PP = "Name Type Size"
    GoTo Tst
Tst:
    Act = LyzObjPP(Obj, PP)
    C
    Return
End Sub

Private Sub ZZ()
Dim A$
Dim B()
Dim C
Dim D%
Dim F$()
Dim XX
End Sub

Sub StopEr(Er$())
If Si(Er) = 0 Then Exit Sub
BrwAy Er
Stop
End Sub

Sub ThwEqObj(A, B, Fun$, Optional Msg$ = "Two given object cannot be same")
If IsEqObj(A, B) Then Thw Fun, Msg
End Sub

