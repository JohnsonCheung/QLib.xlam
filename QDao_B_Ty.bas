Attribute VB_Name = "QDao_B_Ty"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Dic."
Private Const Asm$ = "QDao"
Public Const ShtTySS$ = " A Att B Bool Byt C Chr D Dbl Dte Dec I Int L Lng M Mem S T Tim Txt "
Enum EmSimTy
    EiUnk
    EiEmp
    EiYes
    EiNum
    EiDte
    EiStr
End Enum


Function AyDaoTy(A As Dao.DataTypeEnum)
Dim O
Select Case A
Case Dao.DataTypeEnum.dbBigInt: O = EmpLngAy
End Select
End Function
Function AyDic_RsKF(A As Dao.Recordset, DicKeyFld, AyFld) As Dictionary _
'Return a dictionary of Ay using KeyFld and AyFld.  The Val-of-returned-Dic is Ay using the AyFld.Type to create
Dim O As New Dictionary
Dim K, V
Dim Emp
Dim Ay
    Emp = AyDaoTy(A.Fields(AyFld).Type)
    Ay = Emp
With A
    While Not .EOF
        K = .Fields(DicKeyFld).Value
        V = .Fields(AyFld).Value
        If O.Exists(K) Then
            If True Then
                Ay = O(K)
                PushI Ay, V
                O(K) = Ay
            Else
                PushI O(K), V '<-- It does not work
            End If
        Else
            Ay = Emp
            PushI Ay, V
            O.Add K, Ay
        End If
        .MoveNext
    Wend
End With
Set AyDic_RsKF = O
End Function


Function JnStrDicTwoFldRs(A As Dao.Recordset, Optional Sep$ = " ") As Dictionary
Set JnStrDicTwoFldRs = JnStrDicRsKeyJn(A, 0, 1, Sep)
End Function

Function JnStrDicRsKeyJn(A As Dao.Recordset, KeyFld, JnStrFld, Optional Sep$ = " ") As Dictionary
Dim O As New Dictionary
Dim K, V$
While Not A.EOF
    K = A.Fields(KeyFld).Value
    V = Nz(A.Fields(JnStrFld).Value, "")
    If O.Exists(K) Then
        O(K) = O(K) & Sep & V
    Else
        O.Add K, CStr(Nz(V))
    End If
    A.MoveNext
Wend
Set JnStrDicRsKeyJn = O
End Function

Function DiKqCntzRs(A As Dao.Recordset, Optional Fld = 0) As Dictionary
Set DiKqCntzRs = DiKqCnt(AvRsCol(A))
End Function

Property Get DShtTy() As Drs
Dim Dy(), I
For Each I In SyzSS(ShtTySS)
    PushI Dy, Sy(I, DtaTyzShtTy(I))
Next
DShtTy = DrszFF("ShtTy DtaTy", Dy)
End Property

Property Get ShtTyAy() As String()
ShtTyAy = SyzSS(ShtTySS)
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

Function DtaTyzTF$(D As Database, T, F$)
DtaTyzTF = DtaTy(FdzTF(D, T, F).Type)
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
Case Dao.DataTypeEnum.dbLongBinary: O = "LongBinary"
Case Else: Stop
End Select
DtaTy = O
End Function

Function SimTyzLo(L As ListObject) As EmSimTy()
Dim Sq(): Sq = SqzLo(L)
Dim C%: For C = 1 To UBound(Sq, 2)
    PushI SimTyzLo, SimTyzCol(ColzSq(Sq, C))
Next
End Function
Function SimTy(V) As EmSimTy
SimTy = SimTyzV(VarType(V))
End Function

Function SimTyzV(V As VbVarType) As EmSimTy
Dim O As EmSimTy
Select Case True
Case V = Empty: O = EiEmp
Case V = vbBoolean: O = EiYes
Case V = vbByte, V = vbCurrency, V = vbDecimal, V = vbDouble, V = vbInteger, V = vbLong, V = vbSingle: O = EiNum
Case V = vbDate: O = EiDte
Case V = vbString: O = EiStr
End Select
SimTyzV = O
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

Function DaoTyzV(V) As Dao.DataTypeEnum
Dim T As VbVarType: T = VarType(V)
If T = vbString Then
    If Len(V) > 255 Then
        DaoTyzV = dbMemo
    Else
        DaoTyzV = dbText
    End If
    Exit Function
End If
DaoTyzV = DaoTyzVbTy(T)
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
    If Not IsShtTy(CStr(ShtTy)) Then
        PushI ErzShtTyLis, ShtTy
    End If
Next
End Function

Function IsShtTy(S) As Boolean
Select Case Len(S)
Case 1, 3
    If Not IsAscUCas(Asc(S)) Then Exit Function
    IsShtTy = HasSubStr(ShtTySS, " " & S & " ")
End Select
End Function

Function DtaTyAyzS(ShtTyAy$()) As String()
Dim ShtTy: For Each ShtTy In Itr(ShtTyAy)
    PushI DtaTyAyzS, DtaTyzShtTy(ShtTy)
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
ShtTyAyzShtTyLis = CmlAy(ShtTyLis)
End Function



Function ShtAdoTyAy(A() As AdoDb.DataTypeEnum) As String()
Dim I
For Each I In Itr(A)
    PushI ShtAdoTyAy, ShtAdoTy(CLng(I))
Next
End Function

Function ShtAdoTy$(A As AdoDb.DataTypeEnum)
Dim O$
Select Case A
Case AdoDb.DataTypeEnum.adTinyInt: O = "Byt"
Case AdoDb.DataTypeEnum.adInteger: O = "Lng"
Case AdoDb.DataTypeEnum.adSmallInt: O = "Int"
Case AdoDb.DataTypeEnum.adDate: O = "Dte"
Case AdoDb.DataTypeEnum.adVarChar: O = "Txt"
Case AdoDb.DataTypeEnum.adBoolean: O = "Yes"
Case AdoDb.DataTypeEnum.adDouble: O = "Dbl"
Case AdoDb.DataTypeEnum.adCurrency: O = "Cur"
Case AdoDb.DataTypeEnum.adSingle: O = "Sng"
Case AdoDb.DataTypeEnum.adDecimal: O = "Dec"
Case AdoDb.DataTypeEnum.adVarWChar: O = "Mem"
Case Else: O = "?" & A & "?"
End Select
ShtAdoTy = O
End Function



Function ShtTyDic(FxOrFb$, T) As Dictionary
Select Case True
Case IsFb(FxOrFb): Set ShtTyDic = ShtTyDiczFbt(FxOrFb, T)
Case IsFx(FxOrFb): Set ShtTyDic = ShtTyDiczFxw(FxOrFb, T)
Case Else: Thw CSub, "FxOrFb should be Fx or Fb", "FxOrFb T", FxOrFb, T
End Select
End Function

Private Function ShtTyDiczFbt(Fb, T) As Dictionary
Dim F As Dao.Field
Set ShtTyDiczFbt = New Dictionary
For Each F In Db(Fb).TableDefs(T).Fields
    ShtTyDiczFbt.Add F.Name, ShtTyzDao(F.Type)
Next
End Function

Private Function ShtTyDiczFxw(Fx, W) As Dictionary
Dim C As Column, Cat As Catalog, I
Set Cat = CatzFx(Fx)
For Each I In Cat.Tables(CattnzWsn(W)).Columns
    ShtTyDiczFxw.Add C.Name, ShtTyzAdo(C.Type)
Next
End Function



