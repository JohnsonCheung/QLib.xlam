Attribute VB_Name = "QVb_Ay_Op_AyOp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Ay_Op."
Enum EmIxCol
EiBeg1
EiBeg0
EiBegI
EiNoIx
End Enum
Function DashLT1Ay(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushNDup DashLT1Ay, BefOrAll(I, "_")
Next
End Function

Function AeBlnkEleAtEnd(A$()) As String()
If Si(A) = 0 Then Exit Function
If LasEle(A) <> "" Then AeBlnkEleAtEnd = A: Exit Function
Dim J%
For J = UB(A) To 0 Step -1
    If Trim(A(J)) <> "" Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        AeBlnkEleAtEnd = O
        Exit Function
    End If
Next
End Function

Function ReOrdAy(Ay, SubAy)
Dim HasSubAy: HasSubAy = IntersectAy(Ay, SubAy)
Dim Rest: Rest = MinusAy(Ay, SubAy)
ReOrdAy = AddAy(HasSubAy, Rest)
End Function

Function IntersectAy(A, B)
IntersectAy = ResiU(A)
If Si(A) = 0 Then Exit Function
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If HasEle(B, V) Then PushI IntersectAy, V
Next
End Function
Function MinusAyAp(Ay, ParamArray Ap())
Dim IAy, O
O = Ay
For Each IAy In Ap
    O = MinusAy(Ay, IAy)
    If Si(O) = 0 Then MinusAyAp = O: Exit Function
Next
MinusAyAp = O
End Function

Function MinusSy(A$(), B$()) As String()
MinusSy = MinusAy(A, B)
End Function


Function MinusAy(A, B)
If Si(B) = 0 Then MinusAy = A: Exit Function
MinusAy = ResiU(A)
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If Not HasEle(B, V) Then
        PushI MinusAy, V
    End If
Next
End Function

Function MinEle(Ay)
Dim O: O = FstEle(Ay)
Dim I: For Each I In Itr(Ay)
    If I < O Then O = I
Next
MinEle = O
End Function

Function MaxEle(Ay)
Dim O: O = FstEle(Ay)
Dim I: For Each I In Itr(Ay)
    If I > O Then O = I
Next
MaxEle = O
End Function

Function Ny(S) As String()
Ny = WrdAy(S)
End Function

Function CvVy(Vy)
Const CSub$ = CMod & "CvVy"
Select Case True
Case IsStr(Vy): CvVy = SyzSS(CStr(Vy))
Case IsArray(Vy): CvVy = Vy
Case Else: Thw CSub, "VyzDicKK should either be string or array", "Vy-TypeName Vy", TypeName(Vy), Vy
End Select
End Function

Function CvBytAy(A) As Byte()
CvBytAy = A
End Function

Function CvAv(A) As Variant()
CvAv = A
End Function
Function CvObj(A) As Object
Set CvObj = A
End Function
Function CvEr(A) As VBA.ErrObject

End Function
Function CvSy(A) As String()
Select Case True
Case IsSy(A): CvSy = A
Case IsArray(A): CvSy = SyzAy(A)
Case IsEmpty(A) Or IsMissing(A)
Case Else: CvSy = Sy(A)
End Select
End Function

Function SyShow(XX$, Sy$()) As String()
Dim O$()
Select Case Si(Sy)
Case 0
    Push O, XX & "()"
Case 1
    Push O, XX & "(" & Sy(0) & ")"
Case Else
    Push O, XX & "("
    PushAy O, Sy
    Push O, XX & ")"
End Select
SyShow = O
End Function

Private Sub Z()
Dim A
Dim B()
Dim C$
Dim D$()
Dim XX
CvSy A
Sy B
SyShow C, D
End Sub


Function RmvFstChrzAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI RmvFstChrzAy, RmvFstChr(I)
Next
End Function

Function RmvFstNonLetterzAy(Ay) As String() 'Gen:AyXXX
Dim I
For Each I In Itr(Ay)
    PushI RmvFstNonLetterzAy, RmvFstNonLetter(I)
Next
End Function
Function RmvLasChrzAy(Ay) As String()
'Gen:AyFor RmvLasChr
Dim I
For Each I In Itr(Ay)
    PushI RmvLasChrzAy, RmvLasChr(I)
Next
End Function

Function RmvPfxzAy(Ay, Pfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushI RmvPfxzAy, RmvPfx(I, Pfx)
Next
End Function

Function AeSngQRmk(Ay) As String()
Dim I, S$
For Each I In Itr(Ay)
    S = I
    If Not IsSngQRmk(S) Then PushI AeSngQRmk, S
Next
End Function

Function RmvSngQtezAy(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI RmvSngQtezAy, RmvSngQte(I)
Next
End Function

Function RmvT1zAy(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI RmvT1zAy, RmvT1(I)
Next
End Function

Function RmvTTzAy(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI RmvTTzAy, RmvTT(I)
Next
End Function

Function RplAy(Ay, Fm$, By$, Optional Cnt& = 1) As String()
Dim I
For Each I In Itr(Ay)
    PushI RplAy, Replace(I, Fm, By, Count:=Cnt)
Next
End Function
Function Rmv2DashzAy(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI Rmv2DashzAy, Rmv2Dash(I)
Next
End Function

Function RplStarzAy(Ay, By) As String()
Dim I
For Each I In Itr(Ay)
    PushI RplStarzAy, Replace(I, By, "*")
Next
End Function

Function RplT1zAy(Ay, NewT1) As String()
RplT1zAy = AddPfxzAy(RmvT1zAy(Ay), NewT1 & " ")
End Function

Function OffsetzEmBeg(B As EmIxCol, Optional FmI&) As Byte
Select Case True
Case B = EiBeg0: OffsetzEmBeg = 0
Case B = EiBeg1: OffsetzEmBeg = 1
Case B = EiBegI: OffsetzEmBeg = FmI
Case Else: Thw CSub, "EmIxCol value error", "EmIxCol", B
End Select
End Function

Function AddIxPfx(Ay, Optional B As EmIxCol = EiBeg0, Optional FmI&) As String()
If B = EiNoIx Then AddIxPfx = CvSy(Ay): Exit Function
Dim L, J&, N%
J = OffsetzEmBeg(B, FmI)
N = Len(CStr(UB(Ay) + J))
For Each L In Itr(Ay)
    PushI AddIxPfx, AlignR(J, N) & ": " & L
    J = J + 1
Next
End Function

Function T1Ay(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI T1Ay, T1(I)
Next
End Function

Function T2Ay(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI T2Ay, T2(L)
Next
End Function

Function T3Ay(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI T3Ay, T3(L)
Next
End Function
Function TabN$(N%)
TabN = Space(4 * N)
End Function

Function TabNmV$(Nm$, V, Optional NTab% = 1)
TabNmV = TabN(NTab) & Nm & V
End Function
Function TabNmLy(Nm$, Ly$(), Optional NTab% = 1, Optional Beg As EmIxCol = EiNoIx) As String()
Stop
If Si(Ly) = 0 Then
    PushI TabNmLy, TabN(NTab) & Nm
    Exit Function
End If
Dim Ly1$(), L0$, S$, J&
Ly1 = AddIxPfx(Ly, Beg)
PushI TabNmLy, TabN(NTab) & Nm & Ly1(0)
'
S = TabN(NTab) & Space(Len(Nm))
For J = 1 To UB(Ly1)
    PushI TabNmLy, S & Ly1(J)
Next
End Function
Function TabAy(Ay, Optional NTab% = 1) As String()
TabAy = AddPfxzAy(Ay, TabN(NTab))
End Function

Sub SplitAsgzAyPred(Ay, P As IPred, OTrueAy, OFalseAy)
Dim V: For Each V In Itr(Ay)
    If P.Pred(V) Then
        Push OTrueAy, V
    Else
        Push OFalseAy, V
    End If
Next
End Sub

Function AddSS(Sy$(), SS$) As String()
AddSS = SyzAp(Sy, SyzSS(SS))
End Function

Function LookupT1$(Itm, TkssAy$())
Dim L$, I, Kss$, T1$
For Each I In TkssAy
    L = I
    AsgTRst L, T1, Kss
    If HitKss(Itm, Kss) Then LookupT1 = T1: Exit Function
Next
End Function

