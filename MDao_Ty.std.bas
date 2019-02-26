Attribute VB_Name = "MDao_Ty"
Option Explicit
Const CMod$ = "MDao__Ty."
Public Const VdtShtTyLis$ = "ABytChrDteDecILMSTim"
Property Get MsgzVdtShtTy() As String()
Erase XX
X " Byt"
X " Chr"
X " Dec"
X " Dte"
X " Tim"
X "A   Att"
X "B   Bool"
X "C   Ccy"
X "D   Dbl"
X "I   Int"
X "L   Lng"
X "M   Mem"
X "T   Txt"
Erase XX
End Property
Function IsShtTy(A) As Boolean
Select Case Len(A)
Case 1, 3: If Not IsAscUCase(Asc(A)) Then Exit Function
    IsShtTy = HasSubStr(VdtShtTyLis, A, IgnCas:=True)
End Select
End Function

Function DaoTyzShtTy(ShtTy) As Dao.DataTypeEnum
Dim O As Dao.DataTypeEnum
Select Case ShtTy
Case "A":   O = dbAttachment
Case "B":   O = dbBoolean
Case "Byt": O = dbByte
Case "C":   O = dbCurrency
Case "Chr": O = dbChar
Case "Dte": O = dbDate
Case "Dec": O = dbDecimal
Case "D":   O = dbDouble
Case "I":   O = dbInteger
Case "L":   O = dbLong
Case "M":   O = dbCurrency
Case "S":   O = dbMemo
Case "T":   O = dbText
Case "Tim": O = dbTime
Case Else: Thw CSub, "Invalid ShtTy", "The-Invalid-ShtTy Valid-ShtTy", ShtTy, MsgzVdtShtTy
End Select
DaoTyzShtTy = O
End Function

Function SqlTyzDao$(T As Dao.DataTypeEnum, Optional Sz%, Optional Precious%)
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
Case Else: Thw CSub, "Unsupported DaoTy, cannot covert to ShtTy"
End Select
ShtTyzDao = O
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
    SySsl("Attachment Boolean Byte Currency Date Decimal Double Integer Long Memo Signle Text")
End Select
DaoTyzDtaTy = O
End Function
Function DaoTyzVbTy(A As VbVarType) As Dao.DataTypeEnum
Dim O As Dao.DataTypeEnum
Select Case A
Case VbVarType.vbInteger: O = dbInteger
Case VbVarType.vbLong: O = dbLong
Case VbVarType.vbString: O = dbText
Case VbVarType.vbDate: O = dbDate
Case Else: Thw CSub, "VbTy cannot convert to DaoTy", "VbTy", A
End Select
DaoTyzVbTy = O
End Function

Function DaoTyVal(V) As Dao.DataTypeEnum
DaoTyVal = DaoTyzVbTy(VarType(V))
End Function
Function CvDaoTy(A) As Dao.DataTypeEnum
CvDaoTy = A
End Function
Function DaoTyAyzShtTyLis$(A() As DataTypeEnum)
Dim O$, I
For Each I In A
    O = O & ShtTyzDao(CvDaoTy(I))
Next
DaoTyAyzShtTyLis = O
End Function

Function ShtTyAyErzShtTyLis(ShtTyLis$) As String()
Dim O$(), ShtTy
For Each ShtTy In CmlAy(ShtTyLis)
    If Not IsVdtShtTy(ShtTy) Then
        PushI ShtTyAyErzShtTyLis, ShtTy
    End If
Next
End Function

Function IsVdtShtTy(A) As Boolean
Select Case Len(A)
Case 1, 3: If Not IsAscUCase(Asc(FstChr(A))) Then Exit Function
    IsVdtShtTy = HasSubStr(VdtShtTyLis, A)
End Select
End Function

Function DtaTyAyzShtTyAy(ShtTyAy$()) As String()
Dim ShtTy
For Each ShtTy In Itr(ShtTyAy)
    PushI DtaTyAyzShtTyAy, DtaTyzShtTy(ShtTy)
Next
End Function

Function DtaTyzShtTy$(ShtTy)
DtaTyzShtTy = DtaTy(DaoTyzShtTy(ShtTy))
End Function
Function ShtTyzAdo$(A As ADODB.DataTypeEnum)
Dim O$
Select Case A
Case ADODB.DataTypeEnum.adTinyInt:  O = "Byt"
Case ADODB.DataTypeEnum.adCurrency: O = "C"
Case ADODB.DataTypeEnum.adDecimal:  O = "Dec"
Case ADODB.DataTypeEnum.adDouble:   O = "D"
Case ADODB.DataTypeEnum.adSmallInt: O = "I"
Case ADODB.DataTypeEnum.adInteger:  O = "L"
Case ADODB.DataTypeEnum.adSingle:   O = "S"
Case ADODB.DataTypeEnum.adChar:     O = "Chr"
Case ADODB.DataTypeEnum.adGUID:     O = "G"
Case ADODB.DataTypeEnum.adVarChar:  O = "M"
Case ADODB.DataTypeEnum.adVarWChar: O = "M"
Case ADODB.DataTypeEnum.adLongVarChar: O = "M"
Case ADODB.DataTypeEnum.adBoolean:  O = "B"
Case ADODB.DataTypeEnum.adDate:     O = "Dte"
'Case ADODB.DataTypeEnum.adTime:     O = "Tim"
Case Else
   Thw CSub, "Not supported Case ADODB type", "ADODBTy", A
End Select
ShtTyzAdo = O
End Function

Function ShtTyAyzShtTyLis(ShtTyLis$) As String()
ShtTyAyzShtTyLis = CmlAy(ShtTyLis)
End Function
