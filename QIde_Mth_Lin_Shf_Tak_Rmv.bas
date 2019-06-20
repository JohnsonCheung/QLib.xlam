Attribute VB_Name = "QIde_Mth_Lin_Shf_Tak_Rmv"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_Shf_Tak_Rmv."
Private Const Asm$ = "QIde"
Type SLin
    S As String
    Lin As String
End Type
Function ShfMthTy$(OLin$)
Dim O$: O = TakMthTy(OLin$)
If O = "" Then Exit Function
ShfMthTy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfTermAftAs$(OLin$)
If Not ShfTerm(OLin, "As") Then Exit Function
ShfTermAftAs = ShfT1(OLin)
End Function
Function ShfShtMdy$(OLin$)
ShfShtMdy = ShtMthMdy(ShfMdy(OLin))
End Function
Function ShfShtMthTy$(OLin$)
ShfShtMthTy = ShtMthTy(ShfMthTy(OLin))
End Function
Function ShfShtMthKd$(OLin$)
ShfShtMthKd = ShtMthKdzShtMthTy(ShtMthTy(ShfMthTy(OLin)))
End Function

Function ShfMdy$(OLin$)
Dim O$
O = MthMdy(OLin):
ShfMdy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfKd$(OLin$)
Dim T$: T = TakMthKd(OLin)
If T = "" Then Exit Function
ShfKd = T
OLin = LTrim(RmvPfx(OLin, T))
End Function

Function ShfMthSfx$(OLin$)
ShfMthSfx = ShfChr(OLin, "#!@#$%^&")
End Function
Sub Z_ShfBef()
Dim L$, Sep$, EptL$
GoSub T0
Exit Sub
T0:
    L = "aaa.bbb"
    Sep = "."
    Ept = "aaa"
    EptL = ".bbb"
    GoTo Tst
Tst:
    Act = ShfBef(L, Sep)
    C
    If L <> EptL Then Stop
    Return
End Sub

Function ShfBef$(OLin$, Sep)
Dim P%: P = InStr(OLin, Sep)
If P = 0 Then Exit Function
ShfBef = Left(OLin, P - 1)
OLin = Mid(OLin, P + Len(Sep) - 1)
End Function

Function ShfBefOrAll$(OLin$, Sep$, Optional NoTrim As Boolean)
Dim P%: P = InStr(OLin, Sep)
If P = 0 Then
    If NoTrim Then
        ShfBefOrAll = OLin
    Else
        ShfBefOrAll = Trim(OLin)
    End If
    OLin = ""
    Exit Function
End If
ShfBefOrAll = Bef(OLin, Sep, NoTrim)
OLin = Aft(OLin, Sep, NoTrim)
End Function
Function ShfDotNm$(OLin$)
OLin = LTrim(OLin)
Dim O$: O = TakDotNm(OLin): If O = "" Then Exit Function
ShfDotNm = O
OLin = RmvPfx(OLin, O)
End Function
Function ShfNm$(OLin$)
Dim O$: O = Nm(OLin): If O = "" Then Exit Function
ShfNm = O
OLin = RmvPfx(OLin, O)
End Function

Function ShfRmk$(OLin$)
Dim L$
L = LTrim(OLin)
If FstChr(L) = "'" Then
    ShfRmk = Mid(L, 2)
    OLin = ""
End If
End Function

Function TakMthKd$(Lin)
TakMthKd = PfxSySpc(Lin, MthKdAy)
End Function

Function TakMthTy$(Lin)
TakMthTy = PfxSySpc(Lin, MthTyAy)
End Function

Function RmvMdy$(Lin)
RmvMdy = LTrim(RmvPfxSySpc(Lin, MthMdyAy))
End Function

Function RmvMthTy$(Lin)
RmvMthTy = RmvPfxSySpc(Lin, MthTyAy)
End Function

