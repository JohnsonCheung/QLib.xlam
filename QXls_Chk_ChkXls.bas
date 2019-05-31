Attribute VB_Name = "QXls_Chk_ChkXls"
Option Explicit
Option Compare Text

Function ChkNoWsCd(WsCdn$) As Boolean
If HasWsCd(WsCdn) Then Exit Function
MsgBox RplVBar("No worksheet code||" & WsCdn), vbCritical
ChkNoWsCd = True
End Function

Function ChkNoLo(Ws As Worksheet, LoNm$) As Boolean
If HasLo(Ws, LoNm) Then Exit Function
MsgBox RplVBar("No Lo: " & LoNm & "||In Ws: " & Ws.Name), vbCritical
ChkNoLo = True: Exit Function
End Function

Function ChkLoCCExact(Lo As ListObject, CC$) As Boolean
Dim FF$: FF = JnSpc(FnyzLo(Lo))
If FF = CC Then Exit Function
MsgBox FmtQQ("Expected: ?|Actual: ?", CC, FF), vbCritical, FmtQQ("Lo[?] fields error", Lo.Name)
ChkLoCCExact = True
End Function
Function ChkLoCCAtLeast(Lo As ListObject, CC$) As Boolean

End Function

Function ChkNotHasFfn(Ffn) As Boolean
If HasFfn(Ffn) Then Exit Function
MsgBox FmtQQ("File not file:|?", Ffn), vbCritical
ChkNotHasFfn = True
End Function

