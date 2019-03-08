Attribute VB_Name = "MVb_Fun"
Option Explicit
Const CMod$ = "MVb_Fun."
Public CSub$
Sub Asg(Fm, OTo)
If IsNumeric(Fm) Then
    If Not IsEmpty(OTo) Then
        If IsNull(Fm) Then
            OTo = 0
        Else
            OTo = Val(Fm)
        End If
        Exit Sub
    End If
End If
If IsObject(Fm) Then
    Set OTo = Fm
Else
    If IsNull(Fm) Then
        OTo = ""
    Else
        OTo = Fm
    End If
End If
End Sub

Function InstrN&(S, SubStr, N%)
Dim P&, J%
For J = 1 To N
    P = InStr(P + 1, S, SubStr)
    If P = 0 Then Exit Function
Next
InstrN = P
End Function

Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function

Private Sub Z_InstrN()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InstrN(S, SubStr, N)
Ass Exp = Act
End Sub

Function Max(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
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

Function CanCvLng(A) As Boolean
On Error GoTo X
Dim L&: L = CLng(A)
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

Function NDig%(Length&)
Const CSub$ = CMod & "NDig"
Select Case True
Case 0 > Length: Thw CSub, "Length cannot <0", "Length", Length
Case 10 > Length: NDig = 1
Case 100 > Length: NDig = 2
Case 1000 > Length: NDig = 3
Case 10000 > Length: NDig = 4
Case 100000 > Length: NDig = 5
Case 1000000 > Length: NDig = 6
Case 10000000 > Length: NDig = 7
Case 100000000 > Length: NDig = 8
Case 1000000000 > Length: NDig = 9
Case Else: NDig = 10
End Select
End Function

Sub Vc(A, Optional Fnn$)
Brw A, Fnn, UseVc:=True
End Sub
Sub Brw(A, Optional Fnn$, Optional UseVc As Boolean)
Select Case True
Case IsStr(A): BrwStr A, Fnn, UseVc
Case IsArray(A): BrwAy A, Fnn, UseVc
Case IsAset(A): CvAset(A).Brw Fnn
Case IsDrs(A): BrwDrs CvDrs(A), Fnn:=Fnn, UseVc:=UseVc
Case IsDic(A): BrwDic CvDic(A), UseVc:=UseVc, InclDicValOptTy:=True
Case IsEmpty(A): Debug.Print "Empty"
Case IsNothing(A): Debug.Print "Nothing"
Case Else: Stop
End Select
End Sub

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

Private Sub Z()
Z_InstrN
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
Dim S$: S = FtLines(PjfPj)
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


