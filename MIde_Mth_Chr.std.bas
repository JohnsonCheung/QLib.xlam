Attribute VB_Name = "MIde_Mth_Chr"
Option Explicit
Const MthChrLis$ = "!@#$%^&"

Function IsMthChr(A$) As Boolean
If Len(A) <> 1 Then Exit Function
IsMthChr = HasSubStr(MthChrLis, A)
End Function

Function ArgTyNmTyChr$(MthChr$)
Dim O$
Select Case MthChr
Case "#": O = "Double"
Case "%": O = "Integer"
Case "!": O = "Signle"
Case "@": O = "Currency"
Case "^": O = "LongLong"
Case "$": O = "String"
Case "&": O = "Long"
Case Else: Stop
End Select
ArgTyNmTyChr = O
End Function

Function RmvMthChr$(A)
RmvMthChr = RmvChr(A, MthChrLis)
End Function

Function ShfMthChr$(OLin)
ShfMthChr = ShfChr(OLin, MthChrLis)
End Function
Function MthChr$(Lin)
If IsMthLin(Lin) Then MthChr = TakMthChr(RmvMthNm3(Lin))
End Function

Function TakMthChr$(S)
TakMthChr = TakChr(S, MthChrLis)
End Function
