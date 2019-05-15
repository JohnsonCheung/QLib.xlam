Attribute VB_Name = "QVb_Is_Var"
Option Explicit
Private Const CMod$ = "MVb_Is_Var."
Private Const Asm$ = "QVb"

Function IsAv(V) As Boolean
IsAv = VarType(V) = vbArray + vbVariant
End Function

Function IsAyDic(V As Dictionary) As Boolean
If Not IsSy(V.Keys) Then Exit Function
If Not IsAyOfAy(V.Items) Then Exit Function
IsAyDic = True
End Function

Function IsAyOfAy(V) As Boolean
If Not IsAv(V) Then Exit Function
Dim X
For Each X In Itr(V)
    If Not IsArray(X) Then Exit Function
Next
IsAyOfAy = True
End Function

Function IsBool(V) As Boolean
IsBool = VarType(V) = vbBoolean
End Function

Function IsByt(V) As Boolean
IsByt = VarType(V) = vbByte
End Function

Function IsBytAy(V) As Boolean
IsBytAy = VarType(V) = vbByte + vbArray
End Function

Function IsDic(V) As Boolean
IsDic = TypeName(V) = "Dictionary"
End Function

Function IsDigit(V) As Boolean
IsDigit = "0" <= V And V <= "9"
End Function

Function IsDte(V) As Boolean
IsDte = VarType(V) = vbDate
End Function

Function IsEq(V, B) As Boolean
If Not IsEqTy(V, B) Then Exit Function
Select Case True
Case IsArray(V): IsEq = IsEqAy(V, B)
Case IsDic(V): IsEq = IsEqDic(CvDic(V), CvDic(B))
Case IsObject(V): IsEq = ObjPtr(V) = ObjPtr(B)
Case Else: IsEq = V = B
End Select
End Function

Function IsEqDic(V As Dictionary, B As Dictionary) As Boolean
If V.Count <> B.Count Then Exit Function
If V.Count = 0 Then IsEqDic = True: Exit Function
Dim K1, K2
K1 = QSrt1(V.Keys)
K2 = QSrt1(B.Keys)
If Not IsEqAy(K1, K2) Then Exit Function
Dim K
For Each K In K1
   If B(K) <> V(K) Then Exit Function
Next
IsEqDic = True
End Function

Function IsEqTy(V, B) As Boolean
IsEqTy = VarType(V) = VarType(B)
End Function
Function IsInt(V) As Boolean
IsInt = VarType(V) = vbInteger
End Function
Function IsIntAy(V) As Boolean
IsIntAy = VarType(V) = vbArray + vbInteger
End Function

Function IsItr(V) As Boolean
IsItr = TypeName(V) = "Collection"
End Function

Function IsLetter(V$) As Boolean
Dim C1$: C1 = UCase(V)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsLines(V) As Boolean
If Not IsStr(V) Then Exit Function
IsLines = HasSubStr(V, vbLf)
End Function

Function IsLinesAy(V) As Boolean
If Not IsItrOfSy(V) Then Exit Function
Dim L
For Each L In Itr(V)
    If IsLines(L) Then IsLinesAy = True: Exit Function
Next
End Function

Function IsLng(V) As Boolean
IsLng = VarType(V) = vbLong
End Function

Function IsLngAy(V) As Boolean
IsLngAy = VarType(V) = vbArray + vbLong
End Function

Function IsNe(V, B) As Boolean
IsNe = Not IsEq(V, B)
End Function

Function IsNoLinMd(V As CodeModule) As Boolean
IsNoLinMd = V.CountOfLines = 0
End Function

Function IsNonBlankStr(V) As Boolean
If Not IsStr(V) Then Exit Function
IsNonBlankStr = V <> ""
End Function

Function IsNothing(V) As Boolean
IsNothing = TypeName(V) = "Nothing"
End Function

Function IsObjAy(V) As Boolean
IsObjAy = VarType(V) = vbArray + vbObject
End Function

Function IsPrim(V) As Boolean
Select Case VarType(V)
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

Function IsPun(V$) As Boolean
If IsLetter(V) Then Exit Function
If IsDigit(V) Then Exit Function
If V = "_" Then Exit Function
IsPun = True
End Function

Function IsQuoted(S, Q1$, Optional ByVal Q2$) As Boolean
If Q2 = "" Then Q2 = Q1
If FstChr(S) <> Q1 Then Exit Function
IsQuoted = LasChr(S) = Q2
End Function

Function IsSngQRmk(S) As Boolean
IsSngQRmk = FstChr(LTrim(S)) = "'"
End Function

Function IsSngQuoted(S) As Boolean
IsSngQuoted = IsQuoted(S, "'")
End Function

Function IsSomething(V) As Boolean
IsSomething = Not IsNothing(V)
End Function
Function IsNeedQuote(S) As Boolean
If IsSqBktQuoted(S) Then Exit Function
Select Case True
Case IsAscDig(Asc(FstChr(S))), HasSpc(S), HasDot(S), HasHyphen(S), HasPound(S): IsNeedQuote = True
End Select
End Function
Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function

Function IsStrAy(V) As Boolean
IsStrAy = VarType(V) = vbArray + vbString
End Function
Function IsEmpSy(V) As Boolean
If Not IsSy(V) Then Exit Function
IsEmpSy = Si(V) = 0
End Function
Function IsSy(V) As Boolean
IsSy = IsStrAy(V)
End Function

Function IsTglBtn(V) As Boolean
IsTglBtn = TypeName(V) = "ToggleButton"
End Function


Function IsVbTyNum(V As VbVarType) As Boolean
Select Case V
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

Function IsWhiteChr(V) As Boolean
Select Case Left(V, 1)
Case " ", vbCr, vbLf, vbTab: IsWhiteChr = True
End Select
End Function

Private Sub ZIsSy()
Dim V$()
Dim B: B = V
Dim C()
Dim D
Ass IsSy(V) = True
Ass IsSy(B) = True
Ass IsSy(C) = False
Ass IsSy(D) = False
End Sub

Private Sub ZZ_IsStrAy()
Dim V$()
Dim B: B = V
Dim C()
Dim D
Ass IsStrAy(V) = True
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

Private Sub ZZ()
Z_IsVdtLyDicStr
MVb_IsVar:
End Sub

Function IsAllBlankSy(V$()) As Boolean
Dim I
For Each I In V
    If Trim(I) <> "" Then Exit Function
Next
IsAllBlankSy = True
End Function
Function IsBlankStr(V) As Boolean
If IsStr(V) Then
    If Trim(V) = "" Then IsBlankStr = True
End If
End Function

