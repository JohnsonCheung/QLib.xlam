Attribute VB_Name = "MIde_Mth_PurePrp"
Option Explicit

Sub ImPurePrpPjBrw()
Brw ImpPurePrpLyOfPj
End Sub

Function ImpPurePrpLyOfPj() As String()
ImpPurePrpLyOfPj = ImPurePrpLyzPj(CurPj)
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

Sub PurePrpPjBrw()
Brw PurePrpLyPj
End Sub

Function PurePrpLyPj() As String()
PurePrpLyPj = PurePrpLyzPj(CurPj)
End Function

Function PurePrpLyzPj(A As VBProject) As String()
Dim L
For Each L In Itr(MthLinAyzPj(A))
    If IsPurePrpLin(L) Then PushI PurePrpLyzPj, L
Next
End Function
Function PurePrpIxAy(Src$()) As Long()
Dim Ix&
For Ix = 0 To UB(Src)
    If IsPrpLin(Src(Ix)) Then
        Push PurePrpIxAy, Ix
    End If
Next
End Function

Function PurePrpLnoAy(A As CodeModule) As Long()
Dim O&(), Lno&
For Lno = 1 To A.CountOfLines
    If IsPrpLin(A.Lines(Lno, 1)) Then
        Push O, Lno
    End If
Next
PurePrpLnoAy = O
End Function

Function PurePrpLy(A As CodeModule) As String()
Dim O$(), Lno
For Lno = 0 To Itr(PurePrpLnoAy(A))
    Push O, A.Lines(Lno, 1)
Next
PurePrpLy = O
End Function

Function PurePrpNy(A As CodeModule) As String()
Dim O$(), Lno
For Each Lno In Itr(PurePrpLnoAy(A))
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
If Not HasMthPm(Lin) Then Exit Function
IsImPurePrpLin = Not LetSetPrpNset.Has(MthNm(Lin))
End Function

Function IsPurePrpLin(Lin) As Boolean
If Not MthTy(Lin) = "Property Get" Then Exit Function
If HasMthPm(Lin) Then Exit Function
IsPurePrpLin = True
End Function
Function HasMthPm(MthLin) As Boolean
HasMthPm = BetBkt(MthLin) <> ""
End Function

