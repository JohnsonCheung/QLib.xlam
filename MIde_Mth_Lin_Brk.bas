Attribute VB_Name = "MIde_Mth_Lin_Brk"
Option Explicit
Public Type MthLinRec
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
    .ShtMdy = ShfShtMthMdy(L)
    .ShtTy = ShfShtMthTy(L)
    .Nm = ShfNm(L)
    .TyChr = ShfTyChr(L)
    .Pm = ShfBktStr(L)
    .RetTy = ShfRetTyzAftPm(L)
    .Rmk = RmkzAftRetTy(L)
    .IsRetVal = HasEle(SySsl("Get Fun"), .ShtTy)
    .ShtRetTy = ShtRetTy(.TyChr, .RetTy, .IsRetVal)
End With
End Function

Function MthFLin$(MthQLin)
Dim P$, T$, M$, L$
L = MthQLin
P = ShfTermDot(L)
T = ShfTermDot(L)
M = ShfTermDot(L)
MthFLin = JnDotAp(P, T, M, MthFLinzMthLin(L))
End Function

Function MthFLyOfVbe(Optional WhStr$) As String()
MthFLyOfVbe = MthFLyzVbe(CurVbe, WhStr)
End Function

Function MthFLyzVbe(A As Vbe, Optional WhStr$) As String()
MthFLyzVbe = MthFLy(MthQLyzVbe(A, WhStr))
End Function

Function MthFLy(MthQLy$()) As String()
Dim MthQLin
For Each MthQLin In Itr(MthQLy)
    PushI MthFLy, MthFLin(MthQLin)
Next
End Function

Function MthFLinzMthLin$(MthLin)
Dim X As MthLinRec: X = MthLinRec(MthLin)
With X
Dim RetTy$: RetTy = ShtRetTy(.TyChr, .RetTy, .IsRetVal)
MthFLinzMthLin = JnDotAp(.ShtMdy, .ShtTy, .Nm & RetTy & FmtPm(.Pm)) & IIf(.Rmk = "", "", ".") & .Rmk
End With
End Function

Function FmtPm(Pm$, Optional IsNoBkt As Boolean) 'Pm is wo bkt.
Dim A$: A = Replace(Pm, "Optional ", "?")
Dim B$: B = Replace(A, " As ", ":")
Dim C$: C = Replace(B, "ParamArray ", "...")
If IsNoBkt Then
    FmtPm = C
Else
    FmtPm = QuoteSq(C)
End If
End Function

Function ShtRetTyAsetOfVbe(Optional WhStr$) As Aset
Set ShtRetTyAsetOfVbe = ShtRetTyAsetzVbe(CurVbe, WhStr)
End Function

Function ShtRetTyAsetzVbe(A As Vbe, Optional WhStr$) As Aset
Set ShtRetTyAsetzVbe = ShtRetTyAset(MthLinAyzVbe(A, WhStr))
End Function

Function ShtRetTyAset(MthLinAy$()) As Aset
Set ShtRetTyAset = AsetzAy(ShtRetTyAy(MthLinAy))
End Function

Function ShtRetTyAy(MthLinAy$()) As String()
Dim MthLin
For Each MthLin In Itr(MthLinAy)
    PushI ShtRetTyAy, ShtRetTyzLin(MthLin)
Next
End Function

Function ShtRetTyzLin$(MthLin)
Dim A$: A = MthLinRec(MthLin).ShtRetTy
ShtRetTyzLin = A
If LasChr(A) = ":" Then Stop
End Function

Function ShtRetTy$(TyChr$, RetTy$, IsRetVal As Boolean, Optional ExlColon As Boolean)
Dim O$, Colon$
Colon = IIf(ExlColon, "", ":")
Select Case True
Case Not IsRetVal
Case TyChr = "" And RetTy = "": O = Colon & "Variant"
Case TyChr = "" And RetTy <> "": O = Colon & RetTy
Case RetTy <> "": Thw CSub, "TyChr and RetTy should both have value", "TyChr RetTy", TyChr, RetTy
Case Else: O = TyChr
End Select
ShtRetTy = O
End Function

