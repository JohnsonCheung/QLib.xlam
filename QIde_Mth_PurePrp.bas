Attribute VB_Name = "QIde_Mth_PurePrp"
Option Explicit
Private Const CMod$ = "MIde_Mth_PurePrp."
Private Const Asm$ = "QIde"

Sub ImPurePrpPjBrw()
Brw ImpPurePrpLyP
End Sub

Function ImpPurePrpLyP() As String()
ImpPurePrpLyP = ImPurePrpLyzP(CPj)
End Function

Function ImPurePrpLyzP(P As VBProject) As String()
If IsProtect(P) Then Exit Function
Dim C As VBComponent
For Each C In P.VBComponents
    PushIAy ImPurePrpLyzP, ImPurePrpLyzMd(C.CodeModule)
Next
End Function

Function ImPurePrpLyzMd(A As CodeModule) As String()
ImPurePrpLyzMd = ImPurePrpLyzSrc(Src(A))
End Function

Private Sub Z_ImPurePrpLyzSrc()
Brw ImPurePrpLyzSrc(SrczMdn("MXls_Lo_LofVbl"))
End Sub

Function ImPurePrpLyzSrc(Src$()) As String()
Dim L$, I, M$(), S As New Aset
M = MthLinyzSrc(Src)
Set S = LetSetPrpNset(M)
For Each I In Itr(M)
    L = I
    If IsImPurePrpLin(L, S) Then
        PushI ImPurePrpLyzSrc, L
    End If
Next
End Function

Property Get PurePrpLyP() As String()
PurePrpLyP = PurePrpLyzP(CPj)
End Property

Function PurePrpLyzP(P As VBProject) As String()
Dim L$, I
For Each I In Itr(MthLinyzP(P))
    L = I
    If IsPurePrpLin(L) Then PushI PurePrpLyzP, L
Next
End Function

Function PurePrpLyAyzP(P As VBProject) As Variant()
Dim Ly, C As VBComponent
For Each C In P.VBComponents
    For Each Ly In Itr(PurePrpLyAyzMd(C.Codmodule))
        PushI PurePrpLyAyzP, Ly
    Next
Next
End Function

Function IxyzPurePrp(Src$()) As Long()
Dim Ix&
For Ix = 0 To UB(Src)
    If IsPurePrpLin(Src(Ix)) Then
        Push IxyzPurePrp, Ix
    End If
Next
End Function

Function PurePrpLyAyzMd(A As CodeModule) As Variant()
PurePrpLyAyzMd = PurePrpLyAyzSrc(Src(A))
End Function

Function PurePrpLyAyzSrc(Src$()) As Variant()
Dim Ix&, I
For Each I In Itr(IxyzPurePrp(Src))
    Ix = I
'    PushI PurePrpLyAyzSrc, MthLyBySrcFm(Src, Ix)
Next
End Function
Function PurePrpNy(A As CodeModule) As String()
Dim O$(), Lno
For Each Lno In Itr(IxyzPurePrp(Src(A)))
    PushNoDup O, PrpNm(A.Lines(Lno, 1))
Next
PurePrpNy = O
End Function

Function LetSetPrpNset(MthLiny$()) As Aset
Dim O As New Aset, N$, L$, I
For Each I In Itr(MthLiny)
    L = I
    N = LetSetPrpNm(L)
    'If HasPfx(L, "Property Let") Then Stop
    If N <> "" Then O.PushItm N
Next
Set LetSetPrpNset = O
End Function

Private Function LetSetPrpNm$(Lin)
With Mthn3(Lin)
    Select Case .ShtTy
    Case "Set", "Let": LetSetPrpNm = .Nm: Exit Function
    End Select
End With
End Function
Function IsImPurePrpLin(Lin, LetSetPrpNset As Aset) As Boolean
If Not MthTy(Lin) = "Property Get" Then Exit Function
Stop
If Not HasMthPm(Lin) Then Exit Function
IsImPurePrpLin = Not LetSetPrpNset.Has(Mthn(Lin))
End Function

Function IsPurePrpLin(Lin) As Boolean
Dim O As Boolean
Select Case MthTy(Lin)
Case "Property Get":  O = Not HasMthPm(Lin)
End Select
IsPurePrpLin = O
End Function

Function HasMthPm(MthLin) As Boolean
HasMthPm = MthPm(MthLin) <> ""
End Function

