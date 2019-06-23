Attribute VB_Name = "QXls_Formula_Filter"
Option Explicit
Option Compare Text

Function FilterzLo(Lo As ListObject, Coln)
'Ret : Set filter of all Lo of CWs @
Dim Ws As Worksheet: Set Ws = CWs
Dim C$: C = "Mthn"
Dim LC  As ListColumn:  Set LC = Lo.ListColumns(C)
Dim OFld%:                OFld = LC.Index
Dim Itm():                 Itm = ColzLc(LC)
Dim Patn$:                Patn = "^Ay"
Dim OSel:                 OSel = AwPatn1(Itm, Patn)
Dim ORg As Range:      Set ORg = Lo.Range
ORg.AutoFilter Field:=OFld, Criteria1:=OSel, Operator:=xlFilterValues
End Function
