Attribute VB_Name = "MVb_Dte"
Option Explicit
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
