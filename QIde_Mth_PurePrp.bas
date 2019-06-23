Attribute VB_Name = "QIde_Mth_PurePrp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_PurePrp."
Private Const Asm$ = "QIde"

Function DMthPrpWiPm() As Drs
Dim A As Drs: A = AddColzHasPm(DMthPrp)
DMthPrpWiPm = DwEqExl(A, "HasPm", True)
End Function

Property Get DMthPrpWoPm() As Drs
Dim A As Drs: A = AddColzHasPm(DMthPrp)
DMthPrpWoPm = DwEqExl(A, "HasPm", False)
End Property

Function LetSetPrpNset(MthLinAy$()) As Aset
Dim O As New Aset, N$, L$, I
For Each I In Itr(MthLinAy)
    L = I
    N = LetSetPrpNm(L)
    'If HasPfx(L, "Property Let") Then Stop
    If N <> "" Then O.PushItm N
Next
Set LetSetPrpNset = O
End Function

Private Function LetSetPrpNm$(Lin)
With Mthn3zL(Lin)
    Select Case .ShtTy
    Case "Set", "Let": LetSetPrpNm = .Nm: Exit Function
    End Select
End With
End Function

