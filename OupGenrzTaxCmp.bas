VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OupGenrzTaxCmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Implements IOupGenr
Const CLib$ = "QTaxCmp."
Const CMod$ = CLib & "OupGenrzTaxCmp."
Const A$ = "A"
Sub GenOupTblFmTmpInp(D As Database)
IOupGenr_GenOupTblFmTmpInp D
End Sub
Sub IOupGenr_GenOupTblFmTmpInp(D As Database)
End Sub
