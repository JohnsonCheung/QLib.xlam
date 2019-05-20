Attribute VB_Name = "QDao_Ty_Ado"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Ty_Ado."
Private Const Asm$ = "QDao"
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

