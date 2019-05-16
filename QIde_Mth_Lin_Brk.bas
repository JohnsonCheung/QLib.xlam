Attribute VB_Name = "QIde_Mth_Lin_Brk"
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Brk."
Type MthLinRec
    ShtMdy As String
    ShtTy As String
    Nm As String
    TyChr As String
    RetTy As String
    Pm As String
    Rmk As String
    IsRetVal As Boolean
    ShtRetTy As String
End Type
Private Function ShfRetTyzAftPm$(OAftPm$)
Dim A$: A = ShfTermAftAs(OAftPm)
If LasChr(A) = ":" Then
    ShfRetTyzAftPm = RmvLasChr(A)
    OAftPm = ":" & OAftPm
Else
    ShfRetTyzAftPm = A
End If
End Function
Private Function RmkzAftRetTy$(AftRetTy$)
Select Case True
Case AftRetTy = "", FstChr(AftRetTy) = ":": Exit Function
End Select
Dim L$: L = LTrim(AftRetTy)
If FstChr(L) = "'" Then RmkzAftRetTy = LTrim(RmvFstChr(L)): Exit Function
Thw CSub, "Something wrong in AftRetTy", "AftRetTy", AftRetTy
End Function

Function MthLinRec(MthLin) As MthLinRec
Dim L$: L = MthLin
With MthLinRec
    .ShtMdy = ShfShtMdy(L)
    .ShtTy = ShfShtMthTy(L)
    .Nm = ShfNm(L)
    .TyChr = ShfTyChr(L)
    .Pm = ShfBktStr(L)
    .RetTy = ShfRetTyzAftPm(L)
    .Rmk = RmkzAftRetTy(L)
    .IsRetVal = HasEle(SyzSS("Get Fun"), .ShtTy)
    .ShtRetTy = ShtRetTy(.TyChr, .RetTy, .IsRetVal)
End With
End Function

Function MthFLin(MthQLin)
Dim P$, T$, M$, L$
L = MthQLin
P = ShfTermDot(L)
T = ShfTermDot(L)
M = ShfTermDot(L)
MthFLin = JnDotAp(P, T, M, MthFLinzMthLin(L))
End Function

Function MthFLyInVbe() As String()
MthFLyInVbe = MthFLyzV(CVbe)
End Function

Function MthFLyzV(A As Vbe) As String()
MthFLyzV = MthFLy(MthQLyzV(A))
End Function

Function MthFLy(MthQLy$()) As String()
Dim MthQLin
For Each MthQLin In Itr(MthQLy)
    PushI MthFLy, MthFLin(MthQLin)
Next
End Function

Function MthFLinzMthLin(MthLin)
Dim X As MthLinRec: X = MthLinRec(MthLin)
With X
Dim RetTy$: RetTy = ShtRetTy(.TyChr, .RetTy, .IsRetVal)
MthFLinzMthLin = JnDotAp(.ShtMdy, .ShtTy, .Nm & RetTy & FmtPm(.Pm)) & IIf(.Rmk = "", "", ".") & .Rmk
End With
End Function
