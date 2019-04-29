Attribute VB_Name = "BDs"
Option Explicit
Type Ds: DsNm As String: N As Integer: Ay() As Dt: End Type
Sub AddDt(O As Ds, M As Dt)
If HasDt(O, M.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", O.DsNm, M.DtNm)
PushDt O, M
End Sub
Function DtzDsNm(A As Ds, DtNm$) As Dt
Dim Ay() As Dt, J%
Ay = A.Ay
For J = 0 To A.N - 1
    If Ay(J).DtNm = DtNm Then
        DtzDsNm = Ay(J)
        Exit Function
    End If
Next
Thw CSub, "No such DtNm in Ds", "Such-DtNm DtNy-In-Ds", DtNm, TnyzDs(A)
End Function

Function TnyzDs(A As Ds) As String()
Dim J%, Ay() As Dt
Ay = A.Ay
For J = 0 To A.N - 1
    PushI TnyzDs, Ay(J).DtNm
Next
End Function

Function HasDt(A As Ds, DtNm$) As Boolean
Dim J%, Ay() As Dt
Ay = A.Ay
For J = 0 To A.N - 1
    If Ay(J).DtNm = DtNm Then HasDt = True: Exit Function
Next
End Function

Sub PushDt(O As Ds, M As Dt)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Sub BrwDs(A As Ds, Optional MaxColWdt% = 100, Optional BrkColVbl$, Optional ShwZer As Boolean, Optional HidIxCol As Boolean)
BrwAy FmtDs(A, MaxColWdt, BrkColVbl, ShwZer, HidIxCol)
End Sub

Sub DmpDs(A As Ds)
DmpAy FmtDs(A)
End Sub

Function StrValzDic$(A As Dictionary, K)
StrValzDic = ValzDic(A, K)
End Function

Function ValzDic(A As Dictionary, K)
If A.Exists(K) Then ValzDic = A(K)
End Function
Function FmtDs(A As Ds, Optional MaxColWdt% = 100, Optional BrkColVbl$, Optional ShwZer As Boolean, Optional NoIxCol As Boolean) As String()
PushI FmtDs, "*Ds " & A.DsNm & " " & String(10, "=")
Dim Dic As Dictionary
    Set Dic = DiczVbl(BrkColVbl)
Dim J%, M As Dt, Ay() As Dt
For J = 0 To A.N - 1
    M = Ay(J)
    PushAy FmtDs, FmtDt(M, MaxColWdt, StrValzDic(Dic, M.DtNm), NoIxCol)
Next
End Function
