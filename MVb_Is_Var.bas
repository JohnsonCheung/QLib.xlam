Attribute VB_Name = "MVb_Is_Var"
Option Explicit

Function IsAv(A) As Boolean
IsAv = VarType(A) = vbArray + vbVariant
End Function

Function IsAyDic(A As Dictionary) As Boolean
If Not IsSy(A.Keys) Then Exit Function
If Not IsAyOfAy(A.Items) Then Exit Function
IsAyDic = True
End Function

Function IsAyOfAy(A) As Boolean
If Not IsAv(A) Then Exit Function
Dim X
For Each X In Itr(A)
    If Not IsArray(X) Then Exit Function
Next
IsAyOfAy = True
End Function

Function IsBool(A) As Boolean
IsBool = VarType(A) = vbBoolean
End Function

Function IsByt(A) As Boolean
IsByt = VarType(A) = vbByte
End Function

Function IsBytAy(A) As Boolean
IsBytAy = VarType(A) = vbByte + vbArray
End Function

Function IsDic(A) As Boolean
IsDic = TypeName(A) = "Dictionary"
End Function

Function IsDigit(A) As Boolean
IsDigit = "0" <= A And A <= "9"
End Function

Function IsDte(A) As Boolean
IsDte = VarType(A) = vbDate
End Function

Function IsEq(A, B) As Boolean
If VarType(A) <> VarType(B) Then Exit Function
Select Case True
Case IsArray(A): IsEq = IsEqAy(A, B)
Case IsObject(A): IsEq = ObjPtr(A) = ObjPtr(B)
Case Else: IsEq = A = B
End Select
End Function

Function IsEqDic(A As Dictionary, B As Dictionary) As Boolean
If A.Count <> B.Count Then Exit Function
If A.Count = 0 Then IsEqDic = True: Exit Function
Dim K1, K2
K1 = AyQSrt(A.Keys)
K2 = AyQSrt(B.Keys)
If Not IsEqAy(K1, K2) Then Exit Function
Dim K
For Each K In K1
   If B(K) <> A(K) Then Exit Function
Next
IsEqDic = True
End Function

Function IsEqTy(A, B) As Boolean
IsEqTy = VarType(A) = VarType(B)
End Function

Function IsIntAy(A) As Boolean
IsIntAy = VarType(A) = vbArray + vbInteger
End Function

Function IsItr(A) As Boolean
IsItr = TypeName(A) = "Collection"
End Function

Function IsLetter(A$) As Boolean
Dim C1$: C1 = UCase(A)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsLines(A) As Boolean
If Not IsStr(A) Then Exit Function
IsLines = HasSubStr(A, vbLf)
End Function

Function IsLinesAy(A) As Boolean
If Not IsAllStrAy(A) Then Exit Function
Dim L
For Each L In Itr(A)
    If IsLines(L) Then IsLinesAy = True: Exit Function
Next
End Function

Function IsLng(A) As Boolean
IsLng = VarType(A) = vbLong
End Function

Function IsLngAy(V) As Boolean
IsLngAy = VarType(V) = vbArray + vbLong
End Function

Function IsNe(A, B) As Boolean
IsNe = Not IsEq(A, B)
End Function

Function IsNoLinMd(A As CodeModule) As Boolean
IsNoLinMd = A.CountOfLines = 0
End Function

Function IsNonBlankStr(V) As Boolean
If Not IsStr(V) Then Exit Function
IsNonBlankStr = V <> ""
End Function

Function IsNothing(A) As Boolean
IsNothing = TypeName(A) = "Nothing"
End Function

Function IsObjAy(A) As Boolean
IsObjAy = VarType(A) = vbArray + vbObject
End Function

Function IsPrim(A) As Boolean
Select Case VarType(A)
Case _
   VbVarType.vbBoolean, _
   VbVarType.vbByte, _
   VbVarType.vbCurrency, _
   VbVarType.vbDate, _
   VbVarType.vbDecimal, _
   VbVarType.vbDouble, _
   VbVarType.vbInteger, _
   VbVarType.vbLong, _
   VbVarType.vbSingle, _
   VbVarType.vbString
   IsPrim = True
End Select
End Function

Function IsPun(A$) As Boolean
If IsLetter(A) Then Exit Function
If IsDigit(A) Then Exit Function
If A = "_" Then Exit Function
IsPun = True
End Function

Function IsQuoted(A, Q1$, Optional ByVal Q2$) As Boolean
If Q2 = "" Then Q2 = Q1
If FstChr(A) <> Q1 Then Exit Function
IsQuoted = LasChr(A) = Q2
End Function

Function IsSngQRmk(A) As Boolean
IsSngQRmk = FstChr(LTrim(A)) = "'"
End Function

Function IsSngQuoted(A) As Boolean
IsSngQuoted = IsQuoted(A, "'")
End Function

Function IsSomething(A) As Boolean
IsSomething = Not IsNothing(A)
End Function
Function IsNeedQuote(A) As Boolean
If IsSqBktQuoted(A) Then Exit Function
Select Case True
Case IsAscDig(Asc(FstChr(A))), HasSpc(A), HasDot(A), HasHyphen(A), HasPound(A): IsNeedQuote = True
End Select
End Function
Function IsStr(A) As Boolean
IsStr = VarType(A) = vbString
End Function

Function IsStrAy(A) As Boolean
IsStrAy = VarType(A) = vbArray + vbString
End Function
Function IsEmpSy(A) As Boolean
If Not IsSy(A) Then Exit Function
IsEmpSy = Si(A) = 0
End Function
Function IsSy(A) As Boolean
IsSy = IsStrAy(A)
End Function

Function IsTgl(A) As Boolean
IsTgl = TypeName(A) = "ToggleButton"
End Function


Function IsVbTyNum(A As VbVarType) As Boolean
Select Case A
Case vbInteger, vbLong, vbDouble, vbSingle, vbDouble: IsVbTyNum = True: Exit Function
End Select
End Function

Function IsVdtLyDicStr(LyDicStr$) As Boolean
If Left(LyDicStr, 3) <> "***" Then Exit Function
Dim I, K$(), Key$
For Each I In SplitCrLf(LyDicStr$)
   If Left(I, 3) = "***" Then
       Key = Mid(I, 4)
       If HasEle(K, Key) Then Exit Function
       Push K, Key
   End If
Next
IsVdtLyDicStr = True
End Function

Function IsWhiteChr(A) As Boolean
Select Case Left(A, 1)
Case " ", vbCr, vbLf, vbTab: IsWhiteChr = True
End Select
End Function

Private Sub ZIsSy()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Ass IsSy(A) = True
Ass IsSy(B) = True
Ass IsSy(C) = False
Ass IsSy(D) = False
End Sub

Private Sub ZZ_IsStrAy()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Ass IsStrAy(A) = True
Ass IsStrAy(B) = True
Ass IsStrAy(C) = False
Ass IsStrAy(D) = False
End Sub

Private Sub Z_IsVdtLyDicStr()
Ass IsVdtLyDicStr(LineszVbl("***ksdf|***ksdf1")) = True
Ass IsVdtLyDicStr(LineszVbl("***ksdf|***ksdf")) = False
Ass IsVdtLyDicStr(LineszVbl("**ksdf|***ksdf")) = False
Ass IsVdtLyDicStr(LineszVbl("***")) = True
Ass IsVdtLyDicStr("**") = False
End Sub

Private Sub Z()
Z_IsVdtLyDicStr
MVb_IsVar:
End Sub

Function IsAllBlankSy(A$()) As Boolean
Dim I
For Each I In A
    If Trim(I) <> "" Then Exit Function
Next
IsAllBlankSy = True
End Function


