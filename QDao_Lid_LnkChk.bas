Attribute VB_Name = "QDao_Lid_LnkChk"
Option Explicit
Private Const Asm$ = "QDao"
Private Const CMod$ = "MDao_Lid_LnkChk."
Function ErzLnkTblzTSrcCn(A As Database, T$, S$, Cn$) As String()
Const CSub$ = CMod & "ErzLnkTblzTSrcCn"
On Error GoTo X:
LnkTbl A, T, S, Cn
Exit Function
X: ErzLnkTblzTSrcCn = _
    LyzFunMsgNap(CSub, "Cannot link", "Db Tbl SrcTbl CnStr Er", DbNm(A), T, S, Cn, Err.Description)
End Function

Function ChkFxww(Fx$, Wsnn$, Optional FxKd$ = "Excel file") As String()
Dim W$, I
If Not HasFfn(Fx$) Then ChkFxww = MsgzMisFfn(Fx$, FxKd): Exit Function
For Each I In Ny(Wsnn)
    W = I
    PushIAy ChkFxww, ChkWs(Fx$, W, FxKd)
Next
End Function
Function ChkWs(Fx$, Wsn$, FxKd$) As String()
If HasFxw(Fx$, Wsn) Then Exit Function
Dim M$
M = FmtQQ("? does not have expected worksheet", FxKd)
ChkWs = LyzFunMsgNap(CSub, M, "Folder File Expected-Worksheet Worksheets-in-file", Pth(Fx$), Fn(Fx$), Wsn, Wny(Fx$))
End Function
Function ChkFxw(Fx$, Wsn$, Optional FxKd$ = "Excel file") As String()
Const CSub$ = CMod & "ChkFxw"
ChkFxw = ChkHasFfn(Fx$, FxKd): If Si(ChkFxw) > 0 Then Exit Function
ChkFxw = ChkWs(Fx$, Wsn, FxKd)
End Function
Function ChkLnkWs(A As Database, T$, Fx$, Wsn$, Optional FxKd$ = "Excel file") As String()
Const CSub$ = CMod & "ChkLnkWs"
Dim O$()
    O = ChkFxw(Fx, Wsn, FxKd)
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
