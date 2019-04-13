Attribute VB_Name = "MIde_Mth_PurePrp"
Option Explicit

Sub ImPurePrpPjBrw()
Brw ImpPurePrpLyInPj
End Sub

Function ImpPurePrpLyInPj() As String()
ImpPurePrpLyInPj = ImPurePrpLyzPj(CurPj)
End Function

Function ImPurePrpLyzPj(A As VBProject) As String()
If IsProtect(A) Then Exit Function
Dim C As VBComponent
For Each C In A.VBComponents
    PushIAy ImPurePrpLyzPj, ImPurePrpLyzMd(C.CodeModule)
Next
End Function

Function ImPurePrpLyzMd(A As CodeModule) As String()
ImPurePrpLyzMd = ImPurePrpLyzSrc(Src(A))
End Function

Private Sub Z_ImPurePrpLyzSrc()
Brw ImPurePrpLyzSrc(SrczMdNm("MXls_Lo_LofVbl"))
End Sub

Function ImPurePrpLyzSrc(Src$()) As String()
Dim L, M$(), S As New Aset
M = MthLinAyzSrc(Src)
Set S = LetSetPrpNset(M)
For Each L In Itr(M)
    If IsImPurePrpLin(L, S) Then
        PushI ImPurePrpLyzSrc, L
    End If
Next
End Function

Property Get PurePrpLyInPj() As String()
PurePrpLyInPj = PurePrpLyzPj(CurPj)
End Property

Function PurePrpLyzPj(A As VBProject) As String()
Dim L
For Each L In Itr(MthLinAyzPj(A))
    If IsPurePrpLin(L) Then PushI PurePrpLyzPj, L
Next
End Function

Function PurePrpLyAyzPj(A As VBProject) As Variant()
Dim Ly, C As VBComponent
For Each C In A.VBComponents
    For Each Ly In Itr(PurePrpLyAyzMd(C.Codmodule))
        PushI PurePrpLyAyzPj, Ly
    Next
Next
End Function

Function PurePrpIxAy(Src$()) As Long()
Dim Ix&
For Ix = 0 To UB(Src)
    If IsPurePrpLin(Src(Ix)) Then
        Push PurePrpIxAy, Ix
    End If
Next
End Function

Function PurePrpLyAyzMd(A As CodeModule) As Variant()
PurePrpLyAyzMd = PurePrpLyAyzSrc(Src(A))
End Function

Function PurePrpLyAyzSrc(Src$()) As Variant()
Dim Ix
For Each Ix In Itr(PurePrpIxAy(Src))
    PushI PurePrpLyAyzSrc, MthLyBySrcFm(Src, Ix)
Next
End Function
Function PurePrpNy(A As CodeModule) As String()
Dim O$(), Lno
For Each Lno In Itr(PurePrpIxAy(Src(A)))
    PushNoDup O, PrpNm(A.Lines(Lno, 1))
Next
PurePrpNy = O
End Function

Function LetSetPrpNset(MthLinAy$()) As Aset
Dim O As New Aset, N$, L
For Each L In Itr(MthLinAy)
    N = LetSetPrpNm(L)
    'If HasPfx(L, "Property Let") Then Stop
    If N <> "" Then O.PushItm N
Next
Set LetSetPrpNset = O
End Function

Private Function LetSetPrpNm$(Lin)
With MthNm3(Lin)
    Select Case .ShtTy
    Case "Set", "Let": LetSetPrpNm = .Nm: Exit Function
    End Select
End With
End Function
Function IsImPurePrpLin(Lin, LetSetPrpNset As Aset) As Boolean
If Not MthTy(Lin) = "Property Get" Then Exit Function
Stop
If Not HasMthPm(Lin) Then Exit Function
IsImPurePrpLin = Not LetSetPrpNset.Has(MthNm(Lin))
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

