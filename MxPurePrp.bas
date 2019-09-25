Attribute VB_Name = "MxPurePrp"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxPurePrp."

Function DoPubPrpWiPm() As Drs
Dim A As Drs: A = AddMthColHasPm(DoPubPrp)
DoPubPrpWiPm = DwEQExl(A, "HasPm", True)
End Function

Property Get DoPubPrpWoPm() As Drs
Dim A As Drs: A = AddMthColHasPm(DoPubPrp)
DoPubPrpWoPm = DwEQExl(A, "HasPm", False)
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

Function LetSetPrpNm$(Lin)
With Mthn3zL(Lin)
    Select Case .ShtTy
    Case "Set", "Let": LetSetPrpNm = .Nm: Exit Function
    End Select
End With
End Function
