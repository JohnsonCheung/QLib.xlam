Attribute VB_Name = "MxThwIf"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxThwIf."
Sub ThwIf_ErMsg(Er$(), Fun$, Msg$, ParamArray Nap())
If Si(Er) = 0 Then Exit Sub
Dim Nav(): If UBound(Nap) > 0 Then Nav = Nap
ThwNav Fun, Msg, AddNmV(Nav, "Er", Er)
End Sub

Sub ThwIf_Er(Er$(), Fun$)
If Si(Er) = 0 Then Exit Sub
BrwAy AddSy(LyzFunMsgNap(Fun, ""), Er)
Halt
End Sub

Sub ThwIf_NotStr(A, Fun$)
If IsStr(A) Then Exit Sub
Thw Fun, "Given parameter should be str, but now TypeName=" & TypeName(A)
End Sub

Sub ThwIf_AyabNE(A, B, Optional N1$ = "Exp", Optional N2$ = "Act")
ThwIf_Er ChkEqAy(A, B, N1, N2), CSub
End Sub

Sub ThwIf_NE(A, B, Optional N1$ = "A", Optional N2$ = "B")
Const CSub$ = CMod & "ThwIf_NE"
ThwIf_DifTy A, B, N1, N2
Dim IsLinesA As Boolean, IsLinesB As Boolean
IsLinesA = IsLines(A)
IsLinesB = IsLines(B)
Select Case True
Case IsLinesA Or IsLinesB: If A <> B Then CprLines CStr(A), CStr(B), Hdr:=FmtQQ("Lines [?] [?] not eq.", N1, N2): Stop: Exit Sub
Case IsStr(A):    If A <> B Then CprStr CStr(A), CStr(B), Hdr:=FmtQQ("String [?] [?] not eq.", N1, N2): Stop: Exit Sub
Case IsDic(A):    If Not IsEqDic(CvDic(A), CvDic(B)) Then BrwCprDic CvDic(A), CvDic(B): Stop: Exit Sub
Case IsArray(A):  ThwIf_DifAy A, B, N1, N2
Case IsObject(A): If ObjPtr(A) <> ObjPtr(B) Then Thw CSub, "Two object are diff", FmtQQ("Ty-? Ty-?", N1, N2), TypeName(A), TypeName(B)
Case Else:
    If A <> B Then
        Thw CSub, "A B NE", "A B", A, B
        Exit Sub
    End If
End Select
End Sub


Sub ThwIf_DifAy(AyA, AyB, N1$, N2$)
ThwIf_DifSi AyA, AyB, CSub
ThwIf_DifTy AyA, AyB, N1, N2
Dim J&, A
For Each A In Itr(AyA)
    If Not IsEq(A, AyB(J)) Then
        Dim NN$: NN = FmtQQ("AyTy AySi Dif-At Ay-?-Ele-?-Ty Ay-?-Ele-?-Ty Ay-?-Ele-Val Ay-?-Ele-Val Ay-? Ay-?", N1, J, N2, J, N1, N2, N1, N2)
        Thw CSub, "There is ele in 2 Ay are diff", NN, TypeName(AyA), Si(AyA), J, TypeName(A), TypeName(AyB(J)), A, AyB(J), AyA, AyB
        Exit Sub
    End If
    J = J + 1
Next
End Sub

Sub ThwIf_DifTy(A, B, Optional N1$ = "A", Optional N2$ = "B")
If TypeName(A) = TypeName(B) Then Exit Sub
Dim NN$
NN = FmtQQ("?-TyNm ?-TyNm", N1, N2)
Thw CSub, "Type Diff", NN, TypeName(A), TypeName(B)
End Sub

Sub ThwIf_DifSi(A, B, Fun$)
If Si(A) <> Si(B) Then Thw Fun, "Si-A <> Si-B", "Si-A Si-B", Si(A), Si(B)
End Sub

Sub ThwIf_DifFF(A As Drs, FF$, Fun$)
If JnSpc(A.Fny) <> FF Then Thw Fun, "Drs-FF <> FF", "Drs-FF FF", JnSpc(A.Fny), FF
End Sub

Sub ThwIf_ObjNE(A, B, Fun$, Msg$, Nav())
If IsEqObj(A, B) Then ThwNav Fun, Msg, Nav
End Sub

Sub ThwIf_NoSrt(Ay, Fun$)
If IsSrtdzAy(Ay) Then Thw Fun, "Array should be sorted", "Ay-Ty Ay", TypeName(Ay), Ay
End Sub


Sub ThwIf_Nothing(A, VarNm$, Fun$)
If Not IsNothing(A) Then Exit Sub
Thw Fun, FmtQQ("Given[?] is nothing", VarNm)
End Sub

Sub ThwIf_NotAy(A, Fun$)
If IsArray(A) Then Exit Sub
Thw Fun, "Given parameter should be array, but now TypeName=" & TypeName(A)
End Sub

Sub ThwIf_EqObj(A, B, Fun$, Optional Msg$ = "Two given object cannot be same")
If IsEqObj(A, B) Then Thw Fun, Msg
End Sub


Sub ThwIf_NegEle(Ay, Fun$)
Const CSub$ = CMod & "ThwIf_NEgEle"
Dim O$()
    Dim I, J&: For Each I In Itr(Ay)
        If I < 0 Then
            PushI O, J & ": " & I
            J = J + 1
        End If
    Next
If Si(O) > 0 Then
    Thw CSub, "In [Ay], there are [negative-element (Ix Ele)]", "Ay Neg-Ele", Ay, O
End If
End Sub


