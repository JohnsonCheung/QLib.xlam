Attribute VB_Name = "MXls_Fill"
Option Explicit
Sub FillSeqH(HBar As Range)
Dim Sq()
Sq = SqVbar(HBar.Rows.Count)
RgzResz(HBar, Sq).Value = Sq
End Sub

Sub FillSeqV(Vbar As Range)
Dim Sq()
Sq = SqHBar(Vbar.Rows.Count)
RgzResz(Vbar, Sq).Value = Sq
End Sub

Sub FillWsNy(At As Range)
RgzAyV WsNy(WbzRg(At)), At
End Sub
