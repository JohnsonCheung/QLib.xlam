Attribute VB_Name = "MXls_Fill"
Option Explicit
Sub FillSeqH(HBar As Range)
Dim Sq()
Sq = SqVBar(HBar.Rows.Count)
RgzReSz(HBar, Sq).Value = Sq
End Sub

Sub FillSeqV(VBar As Range)
Dim Sq()
Sq = SqHBar(VBar.Rows.Count)
RgzReSz(VBar, Sq).Value = Sq
End Sub

Sub FillWsNy(At As Range)
RgzAyV WsNy(WbzRg(At)), At
End Sub
