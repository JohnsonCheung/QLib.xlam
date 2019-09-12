Attribute VB_Name = "MxTy"
Option Compare Text
Option Explicit
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxTy."
Public Const ShtTySS$ = " A Att B Bool Byt C Chr D Dbl Dte Dec I Int L Lng M Mem S T Tim Txt "
Enum EmSimTy
    EiUnk
    EiEmp
    EiYes
    EiNum
    EiDte
    EiStr
End Enum

Function AyDaoTy(A As dao.DataTypeEnum)
Dim O
Select Case A
Case dao.DataTypeEnum.dbBigInt: O = EmpLngAy
End Select
End Function

Function AyDic_RsKF(A As dao.Recordset, DicKeyFld, AyFld) As Dictionary _
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

Function CvDaoTy(A) As dao.DataTypeEnum
CvDaoTy = A
End Function

Function DaoTyzCol(Col()) As dao.DataTypeEnum
DaoTyzCol = DaoTyzVbTy(VbTyzCol(Col))
End Function

Function DaoTyzDtaTy(DtaTy$) As dao.DataTypeEnum
Const CSub$ = CMod & "DaoTy"
Dim O
Select Case DtaTy
Case "Attachment": O = dao.DataTypeEnum.dbAttachment
Case "Boolean":    O = dao.DataTypeEnum.dbBoolean
Case "Byte":       O = dao.DataTypeEnum.dbByte
Case "Currency":   O = dao.DataTypeEnum.dbCurrency
Case "Date":       O = dao.DataTypeEnum.dbDate
Case "Decimal":    O = dao.DataTypeEnum.dbDecimal
Case "Double":     O = dao.DataTypeEnum.dbDouble
Case "Integer":    O = dao.DataTypeEnum.dbInteger
Case "Long":       O = dao.DataTypeEnum.dbLong
Case "Memo":       O = dao.DataTypeEnum.dbMemo
Case "Single":     O = dao.DataTypeEnum.dbSingle
Case "Text":       O = dao.DataTypeEnum.dbText
Case Else: Thw CSub, "Invalid ShtTyzDao", "ShtTyzDao Valid", DtaTy, _
    SyzSS("Attachment Boolean Byte Currency Date Decimal Double Integer Long Memo Signle Text")
End Select
DaoTyzDtaTy = O
End Function

Function DaoTyzShtTy(ShtTy) As dao.DataTypeEnum
Dim O As dao.DataTypeEnum
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

Function DaoTyzV(V) As dao.DataTypeEnum
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

Function DaoTyzVbTy(A As VbVarType) As dao.DataTypeEnum
Dim O As dao.DataTypeEnum
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

Function DiKqCntzRs(A As dao.Recordset, Optional Fld = 0) As Dictionary
Set DiKqCntzRs = DiKqCnt(AvRsCol(A))
End Function

Property Get DShtTy() As Drs
Dim Dy(), I
For Each I In SyzSS(ShtTySS)
    PushI Dy, Sy(I, DtaTyzShtTy(I))
Next
DShtTy = DrszFF("ShtTy DtaTy", Dy)
End Property

Function DtaTy$(T As dao.DataTypeEnum)
Dim O$
Select Case T
Case dao.DataTypeEnum.dbAttachment: O = "Attachment"
Case dao.DataTypeEnum.dbBoolean:    O = "Boolean"
Case dao.DataTypeEnum.dbByte:       O = "Byte"
Case dao.DataTypeEnum.dbCurrency:   O = "Currency"
Case dao.DataTypeEnum.dbDate:       O = "Date"
Case dao.DataTypeEnum.dbDecimal:    O = "Decimal"
Case dao.DataTypeEnum.dbDouble:     O = "Double"
Case dao.DataTypeEnum.dbInteger:    O = "Integer"
Case dao.DataTypeEnum.dbLong:       O = "Long"
Case dao.DataTypeEnum.dbMemo:       O = "Memo"
Case dao.DataTypeEnum.dbSingle:     O = "Single"
Case dao.DataTypeEnum.dbText:       O = "Text"
Case dao.DataTypeEnum.dbChar:       O = "Char"
Case dao.DataTypeEnum.dbTime:       O = "Time"
Case dao.DataTypeEnum.dbLongBinary: O = "LongBinary"
Case Else: Stop
End Select
DtaTy = O
End Function

Property Get DtaTyAy() As String()
DtaTyAy = DtaTyAyzS(ShtTyAy)
End Property

Function DtaTyAyzS(ShtTyAy$()) As String()
Dim ShtTy: For Each ShtTy In Itr(ShtTyAy)
    PushI DtaTyAyzS, DtaTyzShtTy(ShtTy)
Next
End Function

Function DtaTyzShtTy$(ShtTy)
DtaTyzShtTy = DtaTy(DaoTyzShtTy(ShtTy))
End Function

Function DtaTyzTF$(D As Database, T, F$)
DtaTyzTF = DtaTy(FdzTF(D, T, F).Type)
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

Function JnStrDicRsKeyJn(A As dao.Recordset, KeyFld, JnStrFld, Optional Sep$ = " ") As Dictionary
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

Function JnStrDicTwoFldRs(A As dao.Recordset, Optional Sep$ = " ") As Dictionary
Set JnStrDicTwoFldRs = JnStrDicRsKeyJn(A, 0, 1, Sep)
End Function

Function MaxSim(A As EmSimTy, B As EmSimTy) As EmSimTy
MaxSim = Max(A, B)
End Function

Function ShtAdoTy$(A As AdoDB.DataTypeEnum)
Dim O$
Select Case A
Case AdoDB.DataTypeEnum.adTinyInt: O = "Byt"
Case AdoDB.DataTypeEnum.adInteger: O = "Lng"
Case AdoDB.DataTypeEnum.adSmallInt: O = "Int"
Case AdoDB.DataTypeEnum.adDate: O = "Dte"
Case AdoDB.DataTypeEnum.adVarChar: O = "Txt"
Case AdoDB.DataTypeEnum.adBoolean: O = "Yes"
Case AdoDB.DataTypeEnum.adDouble: O = "Dbl"
Case AdoDB.DataTypeEnum.adCurrency: O = "Cur"
Case AdoDB.DataTypeEnum.adSingle: O = "Sng"
Case AdoDB.DataTypeEnum.adDecimal: O = "Dec"
Case AdoDB.DataTypeEnum.adVarWChar: O = "Mem"
Case Else: O = "?" & A & "?"
End Select
ShtAdoTy = O
End Function

Function ShtAdoTyAy(A() As AdoDB.DataTypeEnum) As String()
Dim I
For Each I In Itr(A)
    PushI ShtAdoTyAy, ShtAdoTy(CLng(I))
Next
End Function

Property Get ShtTyAy() As String()
ShtTyAy = SyzSS(ShtTySS)
End Property

Function ShtTyAyzShtTyLis(ShtTyLis$) As String()
ShtTyAyzShtTyLis = CmlAy(ShtTyLis)
End Function

Function ShtTyDic(FxOrFb$, T) As Dictionary
Select Case True
Case IsFb(FxOrFb): Set ShtTyDic = ShtTyDiczFbt(FxOrFb, T)
Case IsFx(FxOrFb): Set ShtTyDic = ShtTyDiczFxw(FxOrFb, T)
Case Else: Thw CSub, "FxOrFb should be Fx or Fb", "FxOrFb T", FxOrFb, T
End Select
End Function

Private Function ShtTyDiczFbt(Fb, T) As Dictionary
Dim F As dao.Field
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

Property Get ShtTyDtaTyLy() As String()
Dim O$(), I
For Each I In ShtTyAy
    PushI O, I & " " & DtaTyzShtTy(CStr(I))
Next
ShtTyDtaTyLy = FmtSyz2Term(O)
End Property

Function ShtTyLiszDaoTyAy$(A() As DataTypeEnum)
Dim O$, I
For Each I In A
    O = O & ShtTyzDao(CvDaoTy(I))
Next
ShtTyLiszDaoTyAy = O
End Function

Function ShtTyzAdo$(A As AdoDB.DataTypeEnum)
Dim O$
Select Case A
Case AdoDB.DataTypeEnum.adTinyInt:  O = "Byt"
Case AdoDB.DataTypeEnum.adCurrency: O = "C"
Case AdoDB.DataTypeEnum.adDecimal:  O = "Dec"
Case AdoDB.DataTypeEnum.adDouble:   O = "D"
Case AdoDB.DataTypeEnum.adSmallInt: O = "I"
Case AdoDB.DataTypeEnum.adInteger:  O = "L"
Case AdoDB.DataTypeEnum.adSingle:   O = "S"
Case AdoDB.DataTypeEnum.adChar:     O = "Chr"
Case AdoDB.DataTypeEnum.adGUID:     O = "G"
Case AdoDB.DataTypeEnum.adVarChar:  O = "M"
Case AdoDB.DataTypeEnum.adVarWChar: O = "M"
Case AdoDB.DataTypeEnum.adLongVarChar: O = "M"
Case AdoDB.DataTypeEnum.adBoolean:  O = "B"
Case AdoDB.DataTypeEnum.adDate:     O = "Dte"
'Case ADODB.DataTypeEnum.adTime:     O = "Tim"
Case Else
   Thw CSub, "Not supported Case ADODB type", "ADODBTy", A
End Select
ShtTyzAdo = O
End Function

Function ShtTyzDao$(A As dao.DataTypeEnum)
Dim O$
Select Case A
Case dao.DataTypeEnum.dbAttachment: O = "A"
Case dao.DataTypeEnum.dbBoolean:    O = "B"
Case dao.DataTypeEnum.dbByte:       O = "Byt"
Case dao.DataTypeEnum.dbCurrency:   O = "C"
Case dao.DataTypeEnum.dbChar:       O = "Chr"
Case dao.DataTypeEnum.dbDate:       O = "Dte"
Case dao.DataTypeEnum.dbDecimal:    O = "Dec"
Case dao.DataTypeEnum.dbDouble:     O = "D"
Case dao.DataTypeEnum.dbInteger:    O = "I"
Case dao.DataTypeEnum.dbLong:       O = "L"
Case dao.DataTypeEnum.dbMemo:       O = "Mem"
Case dao.DataTypeEnum.dbSingle:     O = "S"
Case dao.DataTypeEnum.dbText:       O = "T"
Case dao.DataTypeEnum.dbTime:       O = "Tim"
Case Else: Thw CSub, "Unsupported DaoTy, cannot covert to ShtTy", "DaoTy", A
End Select
ShtTyzDao = O
End Function

Function SimTy(V) As EmSimTy
SimTy = SimTyzV(VarType(V))
End Function

Function SimTyzCol(Col()) As EmSimTy
Dim V: For Each V In Itr(Col)
    Dim O As EmSimTy: O = MaxSim(O, SimTy(V))
    If O = EiStr Then SimTyzCol = O: Exit Function
Next
End Function

Function SimTyzLo(L As ListObject) As EmSimTy()
Dim Sq(): Sq = SqzLo(L)
Dim C%: For C = 1 To UBound(Sq, 2)
    PushI SimTyzLo, SimTyzCol(ColzSq(Sq, C))
Next
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

Function SqlTyzDao$(T As dao.DataTypeEnum, Optional Si%, Optional Precious%)
Stop '
End Function

Function VbTyAy(Ay) As VbVarType()
Dim V
For Each V In Ay
    PushI VbTyAy, VarType(V)
Next
End Function

Function VbTyzCol(Col()) As EmSimTy
Stop

End Function
