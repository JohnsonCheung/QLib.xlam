Attribute VB_Name = "MDao_Lid_LnkChk"
Option Explicit
Const CMod$ = "MDao_Lnk_Tbl."
Function ErzLnkTblzTSrcCn(A As Database, T, S$, Cn$) As String()
Const CSub$ = CMod & "ErzLnkTblzTSrcCn"
On Error GoTo X:
LnkTblzTSCn A, T, S, Cn
Exit Function
X: ErzLnkTblzTSrcCn = _
    LyzFunMsgNap(CSub, "Cannot link", "Db Tbl SrcTbl CnStr Er", DbNm(A), T, S, Cn, Err.Description)
End Function

Function ChkFxww(Fx$, Wsnss$, Optional FxKind$ = "Excel file") As String()
Dim W
If Not HasFfn(Fx) Then ChkFxww = MsgzMisFfn(Fx, FxKind): Exit Function
For Each W In TermAy(Wsnss)
    PushIAy ChkFxww, ChkWs(Fx, W, FxKind)
Next
End Function
Function ChkWs(Fx, Wsn, FxKind$) As String()
If HasFxw(Fx, Wsn) Then Exit Function
Dim M$
M = FmtQQ("? does not have expected worksheet", FxKind)
ChkWs = LyzFunMsgNap(CSub, M, "Folder File Expected-Worksheet Worksheets-in-file", Pth(Fx), Fn(Fx), Wsn, WsNyzFx(Fx))
End Function
Function ChkFxw(Fx, Wsn, Optional FxKind$ = "Excel file") As String()
Const CSub$ = CMod & "ChkFxw"
ChkFxw = ChkHasFfn(Fx, FxKind): If Si(ChkFxw) > 0 Then Exit Function
ChkFxw = ChkWs(Fx, Wsn, FxKind)
End Function
Function ChkLnkWs(A As Database, T, Fx, Wsn, Optional FxKind$ = "Excel file") As String()
Const CSub$ = CMod & "ChkLnkWs"
Dim O$()
    O = ChkFxw(Fx, Wsn, FxKind)
    If Si(O) > 0 Then
        ChkLnkWs = O
        Exit Function
    End If
On Error GoTo X
LnkFxw A, T, Fx, Wsn
Exit Function
X: ChkLnkWs = _
    LyzMsgNap("Error in linking Xls file", "Er LnkFx LnkWs ToDb AsTbl", Err.Description, Fx, Wsn, DbNm(A), T)
End Function
