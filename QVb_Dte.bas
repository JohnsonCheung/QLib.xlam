Attribute VB_Name = "QVb_Dte"
Option Explicit
Private Const CMod$ = "MVb_Dte."
Private Const Asm$ = "QVb"
Property Get CurMth() As Byte
CurMth = Month(Now)
End Property

Function NxtMthzM(M As Byte) As Byte
NxtMthzM = IIf(M = 12, 1, M + 1)
End Function

Function PrvMthzM(M As Byte) As Byte
PrvMthzM = IIf(M = 1, 12, M - 1)
End Function

Function FstDteOfMth(A As Date) As Date
FstDteOfMth = DateSerial(Year(A), Month(A), 1)
End Function

Function IsVdtDte(A) As Boolean
On Error Resume Next
IsVdtDte = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function LasDteOfMth(A As Date) As Date
LasDteOfMth = PrvDte(FstDteOfMth(NxtMth(A)))
End Function

Function NxtMth(A As Date) As Date
NxtMth = DateTime.DateAdd("M", 1, A)
End Function
Function IsHHMMDD(S) As Boolean
Select Case True
Case _
    Len(S) <> 6, _
    Not IsHH(Left(S, 2)), _
    Not Is0059(Mid(S, 3, 2)), _
    Not Is0059(Right(S, 2))
Case Else: IsHHMMDD = True
End Select
End Function
Function IsHH(S) As Boolean
Select Case True
Case _
    Len(S) <> 2, _
    Not IsAllDig(S), _
    "00" > S, S > "23"
Case Else: IsHH = True
End Select
End Function
Function Is0059(S) As Boolean
Select Case True
Case _
    Len(S) <> 2, _
    Not IsAllDig(S), _
    "00" > S, S > "59"
Case Else: Is0059 = True
End Select
End Function
Function IsYYYYMMDD(S) As Boolean
If Len(S) <> 8 Then Exit Function
If Not IsYYYY(Left(S, 4)) Then Exit Function
If Not IsMM(Mid(S, 5, 2)) Then Exit Function
If Not IsDD(Right(S, 2)) Then Exit Function
IsYYYYMMDD = True
End Function
Function IsAllDig(S) As Boolean
Dim J%
For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsAllDig = True
End Function
Function IsMM(S$) As Boolean
If Len(S) <> 2 Then Exit Function
If Not IsAllDig(S) Then Exit Function
If S < "00" Then Exit Function
If S > "12" Then Exit Function
IsMM = True
End Function
Function IsYYYY(S$) As Boolean
Select Case True
Case Len(S) <> 4, Not IsAllDig(S), S < "2000"
Case Else: IsYYYY = True
End Select
End Function
Function IsDD(S$) As Boolean
Select Case True
Case Len(S) <> 2, Not IsAllDig(S), S < "00", "31" < S
Case Else: IsDD = True
End Select
End Function

Function PrvDte(A As Date) As Date
PrvDte = DateAdd("D", -1, A)
End Function

Function YYMM$(A As Date)
YYMM = Right(Year(A), 2) & Format(Month(A), "00")
End Function

Function FstDtezYYMM(YYMM) As Date
FstDtezYYMM = DateSerial(Left(YYMM, 2), Mid(YYMM, 3, 2), 1)
End Function

Function IsVdtYYYYMMDD(A) As Boolean
On Error Resume Next
IsVdtYYYYMMDD = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function FstDtezYM(Y As Byte, M As Byte) As Date
FstDtezYM = DateSerial(2000 + Y, M, 1)
End Function

Function LasDtezYM(Y As Byte, M As Byte) As Date
LasDtezYM = NxtMth(FstDtezYM(Y, M))
End Function

Function YofNxtMzYM(Y As Byte, M As Byte) As Byte
YofNxtMzYM = IIf(M = 12, Y + 1, Y)
End Function

Function YofPrvMzYM(Y As Byte, M As Byte) As Byte
YofPrvMzYM = IIf(M = 1, Y - 1, Y)
End Function

Property Get CurY() As Byte
CurY = CurYY - 2000
End Property

Property Get CurYY%()
CurYY = Year(Now)
End Property
