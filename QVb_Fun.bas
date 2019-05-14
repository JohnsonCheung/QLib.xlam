Attribute VB_Name = "QVb_Fun"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fun."
Public CSub$
Sub Swap(OA, OB)
Dim X
X = OA
OA = OB
OB = X
End Sub
Function CUsr$()
CUsr = Environ$("USERNAME")
End Function
Sub Asg(Fm, OTo)
Select Case True
Case IsObject(Fm): Set OTo = Fm
Case Else: OTo = Fm
End Select
End Sub

Sub AssBet(Fun$, V, FmV, ToV)
If FmV > V Then Thw Fun, "FmV > V", "V FmV ToV", V, FmV, ToV
If ToV < V Then Thw Fun, "ToV < V", "V FmV ToV", V, FmV, ToV
End Sub

Function InStrWiIthSubStr&(S, SubStr, Optional Ith% = 1)
Dim P&, J%
If Ith < 1 Then Thw CSub, "Ith cannot be <1", "Ith", Ith
For J = 1 To Ith
    P = InStr(P + 1, SubStr, SubStr)
    If P = 0 Then Exit Function
Next
InStrWiIthSubStr = P
End Function

Function InStrN&(S, SubStr, Optional N% = 1)
InStrN = InStrWiIthSubStr(S, SubStr, N)
End Function

Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function

Private Sub Z_InStrN()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InStrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InStrN(S, SubStr, N)
Ass Exp = Act
End Sub

Function Max(A, B, ParamArray Ap())
Dim Av(), O
Av = Ap
O = IIf(A > B, A, B)
Dim J%
For J = 1 To UB(Av)
   If Av(J) > O Then O = Av(J)
Next
Max = O
End Function

Function MaxVbTy(A As VbVarType, B As VbVarType) As VbVarType
If A = vbString Or B = vbString Then MaxVbTy = A: Exit Function
If A = vbEmpty Then MaxVbTy = B: Exit Function
If B = vbEmpty Then MaxVbTy = A: Exit Function
If A = B Then MaxVbTy = A: Exit Function
Dim AIsNum As Boolean, BIsNum As Boolean
AIsNum = IsVbTyNum(A)
BIsNum = IsVbTyNum(B)
Select Case True
Case A = vbBoolean And BIsNum: MaxVbTy = B
Case AIsNum And B = vbBoolean: MaxVbTy = A
Case A = vbDate Or B = vbDate: MaxVbTy = vbString
Case AIsNum And BIsNum:
    Select Case True
    Case A = vbByte: MaxVbTy = B
    Case B = vbByte: MaxVbTy = A
    Case A = vbInteger: MaxVbTy = B
    Case B = vbInteger: MaxVbTy = A
    Case A = vbLong: MaxVbTy = B
    Case B = vbLong: MaxVbTy = A
    Case A = vbSingle: MaxVbTy = B
    Case B = vbSingle: MaxVbTy = A
    Case A = vbDouble: MaxVbTy = B
    Case B = vbDouble: MaxVbTy = A
    Case A = vbCurrency Or B = vbCurrency: MaxVbTy = A
    Case Else: Stop
    End Select
Case Else: Stop
End Select
End Function

Function CanCvLng(V) As Boolean
On Error GoTo X
Dim L&: L = CLng(V)
CanCvLng = True
X:
End Function

Function Min(ParamArray A())
Dim O, J&, Av()
Av = A
Min = MinAy(Av)
End Function

Sub SndKeys(A$)
DoEvents
SendKeys A, True
End Sub

Function NDig&(N&)
Dim A$: A = N
NDig = Len(A)
End Function

Sub Vc(A, Optional Fnn$)
Brw A, Fnn, UseVc:=True
End Sub
Sub B(A, Optional Fnn$, Optional UseVc As Boolean)
Brw A, Fnn, UseVc
End Sub
Sub Brw(A, Optional Fnn$, Optional UseVc As Boolean)
BrwAy Fmt(A), Fnn, UseVc
End Sub
Function Fmt(A) As String()
Select Case True
Case IsStr(A): Fmt = Sy(A)
Case IsLinesAy(A): Fmt = FmtLinesAy(A)
Case IsArray(A): Fmt = SyzAy(A)
Case IsAset(A):  Fmt = CvAset(A).Sy
Case IsDic(A): Fmt = FmtDic(CvDic(A), InclValTy:=True, AddIx:=True)
Case IsEmpty(A): Fmt = Sy("#Empty")
Case IsNothing(A): Fmt = Sy("#Nothing")
Case Else: Fmt = Sy("#TypeName:" & TypeName(A))
End Select
End Function
Function Mch(Re As RegExp, S) As MatchCollection
Set Mch = Re.Execute(S)
End Function
Function RegExp(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
If Patn = "" Or Patn = "." Then Exit Function
Dim O As New RegExp
With O
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
Set RegExp = O
End Function

Private Sub ZZ()
Z_InStrN
MVb___Fun:
End Sub

Function SumLngAy@(A&())
Dim L, O@
For Each L In Itr(A)
    O = O + L
Next
SumLngAy = O
End Function

Private Sub Z_SumLngAy()
Dim S$: S = LineszFt(PjfP)
Debug.Assert SumLngAy(AscCntAy(S)) = Len(S)
End Sub

Function AscCntAy(S) As Long()
Dim O&(255), A As Byte
Dim J&
For J = 1 To Len(S)
    A = Asc(Mid(S, J, 1))
    O(A) = O(A) + 1
Next
AscCntAy = O
End Function
Function RemBlkSz&(N&, BlkSz%)
RemBlkSz = (N Mod BlkSz)
End Function

Function NBlk&(N&, BlkSz%)
NBlk = ((N - 1) \ BlkSz) + 1
End Function

