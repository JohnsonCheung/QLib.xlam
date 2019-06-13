Attribute VB_Name = "QDao_Ty"
Option Compare Text
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Ty."
Public Const ShtTyss$ = " A Att B Bool Byt C Chr D Dbl Dte Dec I Int L Lng M Mem S T Tim Txt "
Enum EmSimTy
    EmEmp
    EiYes
    EiNbr
    EiDte
    EiStr
End Enum
Property Get DShtTy() As Drs
Dim Dry(), I
For Each I In SyzSS(ShtTyss)
    PushI Dry, Sy(I, DtaTyzShtTy(I))
Next
DShtTy = DrszFF("ShtTy DtaTy", Dry)
End Property

Property Get ShtTyAy() As String()
ShtTyAy = SyzSS(ShtTyss)
End Property

Property Get ShtTyDtaTyLy() As String()
Dim O$(), I
For Each I In ShtTyAy
    PushI O, I & " " & DtaTyzShtTy(CStr(I))
Next
ShtTyDtaTyLy = FmtSyz2Term(O)
End Property

Property Get DtaTyAy() As String()
DtaTyAy = DtaTyAyzS(ShtTyAy)
End Property

Function IsShtTy(S) As Boolean
Select Case Len(S)
Case 1, 3
    If Not IsAscUCas(Asc(S)) Then Exit Function
    IsShtTy = HasSubStr(ShtTyss, " " & S & " ", IgnCas:=True)
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
Case Else: Thw CSub, "Invalid ShtTy", "The-Invalid-ShtTy Valid-ShtTy", ShtTy, ShtTyDtaTyLy
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

Function SimTy(V) As EmSimTy
SimTy = SimTyzV(VarType(V))
End Function

Function SimTyzV(V As VbVarType) As EmSimTy
Dim O As EmSimTy
Select Case True
Case V = vbBoolean: O = EiYes
Case V = vbByte, V = vbCurrency, V = vbDecimal, V = vbDouble, V = vbInteger, V = vbLong, V = vbSingle: O = EiNbr
Case V = vbDate: O = EiDte
Case V = vbString: O = EiStr
Case Else: Thw CSub, "Unsupported VarType", "VarType", V
End Select
End Function
Function MaxSim(A As EmSimTy, B As EmSimTy) As EmSimTy
MaxSim = Max(A, B)
End Function
Function SimTyzCol(Col()) As EmSimTy
Dim V: For Each V In Itr(Col)
    Dim O As EmSimTy: O = MaxSim(O, SimTy(V))
    If O = EiStr Then SimTyzCol = O: Exit Function
Next
End Function
Function VbTyzCol(Col()) As EmSimTy
Stop

End Function
Function VbTyAy(Ay) As VbVarType()
Dim V
For Each V In Ay
    PushI VbTyAy, VarType(V)
Next
End Function
Function DaoTyzCol(Col()) As Dao.DataTypeEnum
DaoTyzCol = DaoTyzVbTy(VbTyzCol(Col))
End Function

Function DaoTyzDtaTy(DtaTy$) As Dao.DataTypeEnum
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
    IsVdtShtTy = HasSubStr(ShtTyss, " " & S & " ")
End Select
End Function

Function DtaTyAyzS(ShtTyAy$()) As String()
Dim ShtTy
For Each ShtTy In Itr(ShtTyAy)
    PushI DtaTyAyzS, DtaTyzShtTy(CStr(ShtTy))
Next
End Function

Function DtaTyzShtTy$(ShtTy)
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

