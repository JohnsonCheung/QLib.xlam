Attribute VB_Name = "QVb_Fs_KF"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fs_Ffn_MisEr."
Type KF
    Kd As String
    Ffn As String
End Type
Type KFs: N As Integer: Ay() As KF: End Type
Function KFwKd_orDie(A As KFs, Kd) As KF
Dim O As KF: O = KFwKd(A, Kd)
If IsEmpKF(O) Then Thw CSub, "Cannot find Kd in KFs", "Kd KFs", Kd, FmtKFs(A)
End Function
Function IsEmpKF(A As KF) As Boolean
If A.Kd = "" And A.Ffn = "" Then IsEmpKF = True
End Function
Sub DmpKFs(A As KFs)
D FmtKFs(A)
End Sub
Function KFwKd(A As KFs, Kd) As KF
Dim J%, Ay() As KF
Ay = A.Ay
For J = 0 To A.N - 1
    If Ay(J).Kd = Kd Then KFwKd = Ay(J): Exit Function
Next
End Function
Sub PushKF(O As KFs, M As KF)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub
Function FmtKFs$(A As KFs)
Dim J%, O$()
For J = 0 To A.N - 1
    PushI O, FmtKF(A.Ay(J))
Next
FmtKFs = JnCrLf(O)
End Function

Function FmtKF$(A As KF)
FmtKF = A.Kd & " " & A.Ffn
End Function

Private Function ErOfDupKd(A As KFs, Fun$) As String()
Dim Dup$(): Dup = AywDup(KdAy(A))
If Si(Dup) Then ErOfDupKd = LyzMsgNap("There is dup Kd in KFs", "Dup-Kd KFs", Dup, FmtKFs(A))
End Function

Private Function DupFfn(A As KFs) As String()
Dim Dup$(): Dup = AywDup(FfnyzKFs(A))
If Si(Dup) Then Thw CSub, "There is dup Ffn in KFs", "Dup-Ffn KFs", Dup, FmtKFs(A)
End Function

Private Function ErOfDupFfn(A As KFs, Fun$) As String()
Dim Dup$(): Dup = DupFfn(A)
If Si(Dup) Then ErOfDupFfn = LyzFunMsgNap(CSub, "There is dup Kd in KFs", "Dup-Kd KFs", Dup, FmtKFs(A))
End Function
Sub BrwKFs(A As KFs)
B FmtKFs(A)
End Sub
Function SampKFs() As KFs
Erase XX
X "MB52 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\MB52 2018-07-30.xls"
X "UOM  C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\sales text.xlsx"
X "ZHT1 C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipCost\Sample\ZHT1.XLSX"
SampKFs = KFs(XX)
Erase XX
End Function
Function KFzzL(Lin) As KF
With BrkSpc(Lin)
    KFzzL.Kd = .S1
    KFzzL.Ffn = .S2
End With
End Function
Function KFszzL(Ly$()) As KFs
Dim L
For Each L In Itr(Ly)
    PushKF KFszzL, KFzzL(L)
Next
End Function
Function KdAy(A As KFs) As String()
Dim J%
For J = 0 To A.N - 1
    PushI KdAy, A.Ay(J).Kd
Next
End Function

Function FfnyzKFs(A As KFs) As String()
Dim J%
For J = 0 To A.N - 1
    PushI FfnyzKFs, A.Ay(J).Ffn
Next
End Function
Function KFszVbl(Vbl$) As KFs
KFszVbl = KFs(LyzVbl(Vbl))
End Function

Function KFs(Ly$()) As KFs
Dim Lin
For Each Lin In Itr(Ly)
    PushKF KFs, KFzLin(Lin)
Next
End Function
Function KFzLin(Lin) As KF
With BrkSpc(Lin)
KFzLin = KF(.S1, .S2)
End With
End Function
Function ErzKFs(A As KFs, Fun$) As String()
Dim E1$(), E2$(), E3$()
E1 = ErOfDupKd(A, Fun)
E2 = ErOfDupFfn(A, Fun)
E3 = ErOfMisFfn(A, Fun)
ErzKFs = Sy(E1, E2, E3)
End Function
Private Function ErOfMisFfn(A As KFs, Fun$) As String()
ErOfMisFfn = MsgzMisKFs(KFszWhMis(A), Fun)
End Function

Function KF(Kd, Ffn) As KF
KF.Kd = Kd
KF.Ffn = Ffn
End Function
Function SngKF(A As KF) As KFs
PushKF SngKF, A
End Function

Function MsgzMisKF(Mis As KF) As String()
Stop
'MsgzMisKF = MsgzMisKFs(SngKF(Mis))
End Function

Function MsgzMisKFs(Mis As KFs, Fun$) As String()
If Mis.N = 0 Then Exit Function
Dim O$(), T$
PushI O, FmtQQ("? file not found", Mis.N) & PpdIf(Fun, "  @")
Dim J%, L$(), R$()
T = Space(4)
For J = 0 To Mis.N - 1
    With Mis.Ay(J)
    PushI L, T & "In Path":
    PushI L, T & "Missing [" & .Kd & "] file"
    PushI L, ""
    PushI R, Pth(.Ffn)
    PushI R, Fn(.Ffn)
    PushI R, ""
    End With
Next
L = AddSfxIfNonBlankzAy(AlignLzAy(L), ": ")
PushIAy O, JnAyab(L, R)
MsgzMisKFs = O
End Function

Function KFszWhMis(A As KFs) As KFs
Dim J%
For J = 0 To A.N - 1
    With A.Ay(J)
        If Not HasFfn(.Ffn) Then PushKF KFszWhMis, A.Ay(J)
    End With
Next
End Function
