Attribute VB_Name = "QXls_Formula_Filter"

Function FilterzLo(Lo As ListObject, Coln)
'Ret : Set filter of all Lo of CWs @
Dim Ws As Worksheet: Set Ws = CWs
Dim C$: C = "Mthn"
Dim Lc  As ListColumn:  Set Lc = Lo.ListColumns(C)
Dim OFld%:                OFld = Lc.Index
Dim Itm():                 Itm = ColzLc(Lc)
Dim Patn$:                Patn = "^Ay"
Dim OSel:                 OSel = AywPatn1(Itm, Patn)
Dim ORg As Range:      Set ORg = Lo.Range
ORg.AutoFilter Field:=OFld, Criteria1:=OSel, Operator:=xlFilterValues
End Function
