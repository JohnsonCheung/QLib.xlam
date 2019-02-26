Attribute VB_Name = "MVb_Thw"
Option Explicit
Const CMod$ = "MVb_Thw."

Sub ThwNegEle(Ay, Fun$)
Const CSub$ = CMod & "ThwNegEle"
Dim I, J&, O$()
For Each I In Itr(Ay)
    If I < 0 Then
        PushI O, J & ": " & I
        J = J + 1
    End If
Next
If Sz(O) > 0 Then
    Thw CSub, "In [Ay], there are [negative-element (Ix Ele)]", Ay, O
End If
End Sub

Sub ThwNESz(A, B, Fun$)
If Sz(A) <> Sz(B) Then Thw Fun, "Sz-A <> Sz-B", "Sz-A Sz-B", Sz(A), Sz(B)
End Sub

Sub ThwNE(A, B)
Const CSub$ = CMod & "ThwNE"
ThwDifTy A, B
Select Case True
Case IsLines(A) Or IsLines(B):  If A <> B Then CmpStr A, B, Hdr:="A B not eq.  They are lines": Stop: Exit Sub
Case IsStr(A):                  If A <> B Then CmpStr A, B, Hdr:="A B not eq.  They are string": Stop: Exit Sub
Case IsDic(A):                  If Not IsEqDic(CvDic(A), CvDic(B)) Then BrwCmpDicAB CvDic(A), CvDic(B): Stop: Exit Sub
Case IsArray(A):                If Not IsEqAy(A, B) Then Thw CSub, "A B NE", "Ty-A Ty-B A B", TypeName(A), TypeName(B), A, B
Case IsObject(A):               If ObjPtr(A) <> ObjPtr(B) Then Thw CSub, "Two object are diff", "A-&-B-TypeName", TypeName(A)
Case Else:
    If A <> B Then
        Thw CSub, "A B NE", "A B", A, B
        Exit Sub
    End If
End Select
End Sub

Sub ThwDifTy(A, B)
'B) Then XDmp_Lin_Stop Fun, "Type Diff", "A-Ty B-Ty A B", TypeName(A), TypeName(B), A, B: Exit Sub
End Sub

Sub ThwDifSz(A, B, Fun$)
If Sz(A) = Sz(B) Then Exit Sub
Thw Fun, "Two ay has dif sz", "SzA SzB TyA TyB AyA AyB", Sz(A), Sz(B), TypeName(A), TypeName(B), A, B
End Sub

Sub ThwNotExistFfn(Ffn$, Fun$, Optional FilKd$ = "file")
Stop '
End Sub

Sub ThwEqObjNav(A, B, Fun$, Msg$, Nav())
If IsEqObj(A, B) Then ThwNav Fun, Msg, Nav
End Sub

Sub ThwAyNotSrt(Ay, Fun$)
If IsSrtAy(Ay) Then Thw Fun, "Array should be sorted", "Ay-Ty Ay", TypeName(Ay), Ay
End Sub
Sub ThwOpt(Thw As eThwOpt, Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
Select Case Thw
Case eNoThwInfo: InfoNav Fun, Msg, Nav
Case eNoThwNoInfo:
Case Else:   ThwNav Fun, Msg, Nav
End Select
End Sub
Sub Thw(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
ThwNav Fun, Msg, Nav
End Sub

Sub ThwNav(Fun$, Msg$, Nav())
BrwAy LyzFunMsgNav(Fun, Msg, Nav)
Halt
End Sub

Sub Ass(A As Boolean)
Debug.Assert A
End Sub
Sub ThwNothing(A, Fun$)
If IsNothing(A) Then Exit Sub
Thw Fun, "Given parameter should be array, but now TypeName=" & TypeName(A)
End Sub
Sub ThwNotAy(A, Fun$)
If IsArray(A) Then Exit Sub
Thw Fun, "Given parameter should be array, but now TypeName=" & TypeName(A)
End Sub
Sub ThwNever(Fun$, Optional Msg$ = "Program should not reach here")
Thw Fun, Msg
End Sub

Sub Halt(Optional Fun$)
Err.Raise -1, Fun, "Please check messages opened in notepad"
End Sub

Sub Done()
MsgBox "Done"
End Sub
Sub ThwPgmEr(Er$(), Fun$)
If Sz(Er) = 0 Then Exit Sub
BrwAy AyAdd(Box("Programm Error"), Er)
Halt
End Sub
Function NavAddNNAv(Nav(), NN$, Av()) As Variant()
Dim O(): O = Nav
If Sz(O) = 0 Then
    PushI O, NN
Else
    O(0) = O(0) & " " & NN
End If
PushAy O, Av
NavAddNNAv = O
End Function
Function NavAddNmV(Nav(), Nm$, V) As Variant()
NavAddNmV = NavAddNNAv(Nav, Nm, Av(V))
End Function
Sub ThwErMsg(Er$(), Fun$, Msg$, ParamArray Nap())
If Sz(Er) = 0 Then Exit Sub
Dim Nav(): Nav = Nap
ThwNav Fun, Msg, NavAddNmV(Nav, "Er", Er)
End Sub
Sub ThwEr(Er$(), Fun$)
If Sz(Er) = 0 Then Exit Sub
ThwNav Fun, "There is error", Av("Er", Er)
End Sub
Sub ThwLoopingTooMuch(Fun$)
Thw Fun, "Looping too much"
End Sub
Sub ThwPmEr(PmVal, Fun$, Optional MsgWhyPmEr$ = "Invalid value")
Thw Fun, "Parameter error: " & MsgWhyPmEr, "Pm-Type Pm-Val", TypeName(PmVal), LyzVal(PmVal)
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

Sub DmpAsc(S)
Dim J&, C$
Debug.Print "Len=" & Len(S)
For J = 1 To Len(S)
    C = Mid(S, J, 1)
    Debug.Print J, Asc(C), C
Next
End Sub
Sub DmpAy(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print Ay(J)
Next
End Sub

Sub InfoLin(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
D LinzFunMsgNav(Fun, Msg, Nav)
End Sub
Sub InfoNav(Fun$, Msg$, Nav())
D LyzFunMsgNav(Fun, Msg, Nav)
End Sub
Sub Info(Fun$, Msg$, ParamArray Nap())
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

Private Sub Z_InfoObjPP()
Dim Fun$, Msg$, Obj, PP$
Fun = "XXX"
Msg = "MsgABC"
Set Obj = New Dao.Field
PP = "Name Type Size"
GoSub Tst
Exit Sub
Tst:
    InfoObjPP Fun, Msg, Obj, PP
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
If Sz(Er) = 0 Then Exit Sub
BrwAy Er
Stop
End Sub

