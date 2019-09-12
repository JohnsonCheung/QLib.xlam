Attribute VB_Name = "MxMdIdOp"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxMdIdOp."
Const MdIdFn$ = "MdId.csv"

Function MdIdFcsv$()
MdIdFcsv = ResFfn(MdIdFn)
End Function

Sub EdtMdId()
OpnFcsv MdIdFcsv
End Sub

Sub RfhMdIdFcsv()
EnsFt MdIdFcsv, CsvStrzDrs(DoMdIdP)
End Sub

Sub UpdCNsvCLibvCModv()
'Do : Upd-CNs-CLib-CMod ! Upd Const-CNs-CLib-CMod$ & Const-CNs$ from :MdIdFcsv
Dim D As Drs: D = DrszFcsv(MdIdFcsv)
Dim D1 As Drs: D1 = SelDrs(D, "Mdn CNsv CLibv CModv")
Dim P As VBProject: Set P = CPj
Dim C As VBComponent: For Each C In P.VBComponents
    
Next
End Sub
