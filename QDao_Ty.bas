Attribute VB_Name = "QDao_Ty"
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Ty."
Public Const ShtTyLis$ = "AAttBBoolBytCChrDDteDecIIntLLngMMemSTTimTxt"

Property Get ShtTyDrs() As Drs
Dim Dry(), I
For Each I In CmlSy(ShtTyLis)
    PushI Dry, Sy(I, DtaTyzShtTy(CStr(I)))
Next
ShtTyDrs = DrszFF("ShtTy DtaTy", Dry)
End Property

Property Get ShtTySy() As String()
ShtTySy = CmlSy(ShtTyLis)
End Property

Property Get ShtTyDtaTyLy() As String()
Dim O$(), I
For Each I In ShtTySy
    PushI O, I & " " & DtaTyzShtTy(CStr(I))
Next
ShtTyDtaTyLy = FmtSyz2Term(O)
End Property

Property Get DtaTySy() As String()
DtaTySy = DtaTySyzShtTySy(ShtTySy)
End Property

Function IsShtTy(S) As Boolean
Select Case Len(S)
Case 1, 3: If Not IsAscUCas(Asc(S)) Then Exit Function
    IsShtTy = HasSubStr(ShtTyLis, S, IgnCas:=True)
End Select
End Function

Function DaoTyzShtTy(ShtTy$) As Dao.DataTypeEnum
Dim O As Dao.DataTypeEnum
Select Case ShtTy
Case "A", "Att":  O = dbAttachment
Case "B", "Bool":  O = dbBoolean
Case "Byt": O = dbByte
Case "C", "Cur":  O = dbCurrency
Case "Chr": O = dbChar
Case "Dte": O = dbDate
Case "Dec": O = dbDecimal
Case "D", "Dbl":  O = dbDouble
Case "I", "Int":  O = dbInteger
Case "L", "Lng":  O = dbLong
Case "M", "Mem":  O = dbMemo
Case "S", "Sng":  O = dbSingle
Case "T", "Txt":  O = dbText
Case "Tim": O = dbTime
Case Else: ThwShtTyEr CSub, ShtTy
End Select
DaoTyzShtTy = O
End Function

Function SqlTyzDao$(T As Dao.DataTypeEnum, Optional Si%, Optional Precious%)
Stop '
End Function

Function ShtTyzDao$(A As Dao.DataTypeEnum)
Dim O$
Select Case A
Case Dao.DataTypeEnum.dbAttachment: O = "A"
Case Dao.DataTypeEnum.dbBoolean:    O = "B"
Case Dao.DataTypeEnum.dbByte:       O = "Byt"
Case Dao.DataTypeEnum.dbCurrency:   O = "C"
Case Dao.DataTypeEnum.dbChar:       O = "Chr"
Case Dao.DataTypeEnum.dbDate:       O = "Dte"
Case Dao.DataTypeEnum.dbDecimal:    O = "Dec"
Case Dao.DataTypeEnum.dbDouble:     O = "D"
Case Dao.DataTypeEnum.dbInteger:    O = "I"
Case Dao.DataTypeEnum.dbLong:       O = "L"
Case Dao.DataTypeEnum.dbMemo:       O = "Mem"
Case Dao.DataTypeEnum.dbSingle:     O = "S"
Case Dao.DataTypeEnum.dbText:       O = "T"
Case Dao.DataTypeEnum.dbTime:       O = "Tim"
Case Else: Thw CSub, "Unsupported DaoTy, cannot covert to ShtTy", "DaoTy", A
End Select
ShtTyzDao = O
End Function

Function DtaTyzTF$(A As Database, T, F$)
DtaTyzTF = DtaTy(FdzTF(A, T, F).Type)
End Function

Function DtaTy$(T As Dao.DataTypeEnum)
Dim O$
Select Case T
Case Dao.DataTypeEnum.dbAttachment: O = "Attachment"
Case Dao.DataTypeEnum.dbBoolean:    O = "Boolean"
Case Dao.DataTypeEnum.dbByte:       O = "Byte"
Case Dao.DataTypeEnum.dbCurrency:   O = "Currency"
Case Dao.DataTypeEnum.dbDate:       O = "Date"
Case Dao.DataTypeEnum.dbDecimal:    O = "Decimal"
Case Dao.DataTypeEnum.dbDouble:     O = "Double"
Case Dao.DataTypeEnum.dbInteger:    O = "Integer"
Case Dao.DataTypeEnum.dbLong:       O = "Long"
Case Dao.DataTypeEnum.dbMemo:       O = "Memo"
Case Dao.DataTypeEnum.dbSingle:     O = "Single"
Case Dao.DataTypeEnum.dbText:       O = "Text"
Case Dao.DataTypeEnum.dbChar:       O = "Char"
Case Dao.DataTypeEnum.dbTime:       O = "Time"
Case Else: Stop
End Select
DtaTy = O
End Function

Function DaoTyzDtaTy(DtaTy) As Dao.DataTypeEnum
Const CSub$ = CMod & "DaoTy"
Dim O
Select Case DtaTy
Case "Attachment": O = Dao.DataTypeEnum.dbAttachment
Case "Boolean":    O = Dao.DataTypeEnum.dbBoolean
Case "Byte":       O = Dao.DataTypeEnum.dbByte
Case "Currency":   O = Dao.DataTypeEnum.dbCurrency
Case "Date":       O = Dao.DataTypeEnum.dbDate
Case "Decimal":    O = Dao.DataTypeEnum.dbDecimal
Case "Double":     O = Dao.DataTypeEnum.dbDouble
Case "Integer":    O = Dao.DataTypeEnum.dbInteger
Case "Long":       O = Dao.DataTypeEnum.dbLong
Case "Memo":       O = Dao.DataTypeEnum.dbMemo
Case "Single":     O = Dao.DataTypeEnum.dbSingle
Case "Text":       O = Dao.DataTypeEnum.dbText
Case Else: Thw CSub, "Invalid ShtTyzDao", "ShtTyzDao Valid", DtaTy, _
    SyzSS("Attachment Boolean Byte Currency Date Decimal Double Integer Long Memo Signle Text")
End Select
DaoTyzDtaTy = O
End Function

Function DaoTyzVbTy(A As VbVarType) As Dao.DataTypeEnum
Dim O As Dao.DataTypeEnum
Select Case A
Case vbBoolean: O = dbBoolean
Case vbByte: O = dbByte
Case VbVarType.vbCurrency: O = dbCurrency
Case VbVarType.vbDate: O = dbDate
Case VbVarType.vbDecimal: O = dbDecimal
Case VbVarType.vbDouble: O = dbDouble
Case VbVarType.vbInteger: O = dbInteger
Case VbVarType.vbLong: O = dbLong
Case VbVarType.vbSingle: O = dbSingle
Case VbVarType.vbString: O = dbText
Case Else: Thw CSub, "VbTy cannot convert to DaoTy", "VbTy", A
End Select
DaoTyzVbTy = O
End Function

Function DaoTyzVal(V) As Dao.DataTypeEnum
Dim T As VbVarType: T = VarType(V)
If T = vbString Then
    If Len(V) > 255 Then
        DaoTyzVal = dbMemo
    Else
        DaoTyzVal = dbText
    End If
    Exit Function
End If
DaoTyzVal = DaoTyzVbTy(T)
End Function

Function CvDaoTy(A) As Dao.DataTypeEnum
CvDaoTy = A
End Function

Function ShtTyLiszDaoTyAy$(A() As DataTypeEnum)
Dim O$, I
For Each I In A
    O = O & ShtTyzDao(CvDaoTy(I))
Next
ShtTyLiszDaoTyAy = O
End Function

Function ErzShtTyLis(ShtTyLis$) As String()
Dim O$(), ShtTy
For Each ShtTy In CmlSy(ShtTyLis)
    If Not IsVdtShtTy(CStr(ShtTy)) Then
        PushI ErzShtTyLis, ShtTy
    End If
Next
End Function

Function IsVdtShtTy(S) As Boolean
Select Case Len(S)
Case 1, 3: If Not IsAscUCas(Asc(FstChr(S))) Then Exit Function
    IsVdtShtTy = HasSubStr(ShtTyLis, S)
End Select
End Function

Function DtaTySyzShtTySy(ShtTySy$()) As String()
Dim ShtTy
For Each ShtTy In Itr(ShtTySy)
    PushI DtaTySyzShtTySy, DtaTyzShtTy(CStr(ShtTy))
Next
End Function

Function DtaTyzShtTy$(ShtTy$)
DtaTyzShtTy = DtaTy(DaoTyzShtTy(ShtTy))
End Function

Function ShtTyzAdo$(A As AdoDb.DataTypeEnum)
Dim O$
Select Case A
Case AdoDb.DataTypeEnum.adTinyInt:  O = "Byt"
Case AdoDb.DataTypeEnum.adCurrency: O = "C"
Case AdoDb.DataTypeEnum.adDecimal:  O = "Dec"
Case AdoDb.DataTypeEnum.adDouble:   O = "D"
Case AdoDb.DataTypeEnum.adSmallInt: O = "I"
Case AdoDb.DataTypeEnum.adInteger:  O = "L"
Case AdoDb.DataTypeEnum.adSingle:   O = "S"
Case AdoDb.DataTypeEnum.adChar:     O = "Chr"
Case AdoDb.DataTypeEnum.adGUID:     O = "G"
Case AdoDb.DataTypeEnum.adVarChar:  O = "M"
Case AdoDb.DataTypeEnum.adVarWChar: O = "M"
Case AdoDb.DataTypeEnum.adLongVarChar: O = "M"
Case AdoDb.DataTypeEnum.adBoolean:  O = "B"
Case AdoDb.DataTypeEnum.adDate:     O = "Dte"
'Case ADODB.DataTypeEnum.adTime:     O = "Tim"
Case Else
   Thw CSub, "Not supported Case ADODB type", "ADODBTy", A
End Select
ShtTyzAdo = O
End Function

Function ShtTyAyzShtTyLis(ShtTyLis$) As String()
ShtTyAyzShtTyLis = CmlSy(ShtTyLis)
End Function

Sub ThwShtTyEr(Fun$, ShtTy$)
Thw Fun, "Invalid ShtTy", "The-Invalid-ShtTy Valid-ShtTy", ShtTy, ShtTyDtaTyLy
End Sub

