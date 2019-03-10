Attribute VB_Name = "MDao_Ty"
Option Explicit
Const CMod$ = "MDao__Ty."
Public Const ShtTyLis$ = "AAttBBoolBytCChrDDteDecIIntLLngMMemSTTimTxt"

Property Get ShtTyDrs() As Drs
Dim Dry(), I
For Each I In CmlAy(ShtTyLis)
    PushI Dry, Sy(I, DtaTyzShtTy(I))
Next
Set ShtTyDrs = Drs("ShtTy DtaTy", Dry)
End Property

Property Get VdtShtTyAy() As String()
VdtShtTyAy = CmlAy(ShtTyLis)
End Property

Property Get VdtShtTyDtaTyAy() As String()
Dim O$(), I
For Each I In VdtShtTyAy
    PushI O, I & " " & DtaTyzShtTy(I)
Next
VdtShtTyDtaTyAy = FmtAy2T(O)
End Property

Property Get VdtDtaTyAy() As String()
VdtDtaTyAy = DtaTyAyzShtTyAy(VdtShtTyAy)
End Property

Function IsShtTy(A) As Boolean
Select Case Len(A)
Case 1, 3: If Not IsAscUCase(Asc(A)) Then Exit Function
    IsShtTy = HasSubStr(ShtTyLis, A, IgnCas:=True)
End Select
End Function

Function DaoTyzShtTy(ShtTy) As Dao.DataTypeEnum
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
Case Else: Thw CSub, "Unsupported DaoTy, cannot covert to ShtTy", "DaoTy", A
End Select
ShtTyzDao = O
End Function

Function DtaTyzTF$(A As Database, T, F)
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
    SySsl("Attachment Boolean Byte Currency Date Decimal Double Integer Long Memo Signle Text")
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
For Each ShtTy In CmlAy(ShtTyLis)
    If Not IsVdtShtTy(ShtTy) Then
        PushI ErzShtTyLis, ShtTy
    End If
Next
End Function

Function IsVdtShtTy(A) As Boolean
Select Case Len(A)
Case 1, 3: If Not IsAscUCase(Asc(FstChr(A))) Then Exit Function
    IsVdtShtTy = HasSubStr(ShtTyLis, A)
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

Sub ThwShtTyEr(Fun$, ShtTy)
Thw Fun, "Invalid ShtTy", "The-Invalid-ShtTy Valid-ShtTy", ShtTy, VdtShtTyDtaTyAy
End Sub

