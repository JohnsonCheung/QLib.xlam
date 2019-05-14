Attribute VB_Name = "QXls_Fill"
Option Explicit
Private Const CMod$ = "MXls_Fill."
Private Const Asm$ = "QXls"
Sub FillSeqH(HBar As Range)
Dim Sq()
Sq = SqVzN(HBar.Rows.Count)
ResiRg(HBar, Sq).Value = Sq
End Sub

Sub FillSeqV(Vbar As Range)
Dim Sq()
Sq = SqHzN(Vbar.Rows.Count)
ResiRg(Vbar, Sq).Value = Sq
End Sub

Sub FillWsny(At As Range)
RgzAyV Wsny(WbzRg(At)), At
End Sub
