Attribute VB_Name = "QIde_Mth_Pm"
Option Explicit
Private Const CMod$ = "MIde_Mth_Pm_Arg."
Private Const Asm$ = "QIde"
Public Const DoczShtArg$ = "It is string from Arg"
Public Const DoczArg$ = "It is Sy.  It is splitting of MthPm"
Public Const DoczArgSy$ = "It Array of Arg"
Public Const DoczSset$ = "String-Aset"
Public Const DoczArgTy$ = "It is a string defining the type of an arg.  Eg, Dim A() as integer => ArgTy[Integer()].  Dim A%() => ArgTy[%()]"
Function MthPm$(MthLin)
MthPm = BetBktMust(MthLin, CSub)
End Function

Property Get ArgAsetP() As Aset
Set ArgAsetP = ArgAsetzP(CPj)
End Property

Function ArgAsetzP(P As VBProject) As Aset
Set ArgAsetzP = New Aset
Dim L$, I
'For Each I In MthLinyzP(A)
    L = I
    'ArgAsetzPj.PushAy ArgSy(L)
'Next
End Function

Private Sub Z_ArgAsetP()
ArgAsetP.Srt.Vc
End Sub

Function DimItmzArg$(Arg$)
DimItmzArg = BefOrAll(RmvPfxSpc(RmvPfxSpc(Arg, "Optional"), "ParamArray"), " =")
End Function

Function ShfArgPfx$(OLin$)
Select Case True
Case ShfTerm(OLin, "Optional "):   ShfArgPfx = "?"
Case ShfTerm(OLin, "Paramarray "): ShfArgPfx = ".."
End Select
End Function
Function ArgTy$(AftNm$)

End Function
Function ShtArg$(Arg$)
Dim L$:     L = Arg
Dim Pfx$:     Pfx = ShfArgPfx(L)
Dim Ty$: Ty = ArgTy(L)
'ShtArg = Pfx & Nm & Ty
End Function
