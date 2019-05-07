Attribute VB_Name = "QIde_Rf_RfOp"
Sub AddRfzRff(A As VBProject, Frfee$)
Const CSub$ = CMod & "AddRf"
If HasFrfee(A, Frfee) Then
    InfLin CSub, "Frfee exists in Pj", "Frfee Pj", Frfee, A.Name
    Exit Sub
End If
A.References.AddFromFile Frfee
InfLin CSub, "Frfee is added to Pj", "Frfee Pj", Frfee, A.Name
End Sub

Sub AddRfzSrc(A As VBProject, RfSrc$())
Dim I
For Each I In Itr(RfSrc)
    AddRf A, RfLin(CStr(I))
Next
End Sub

Sub AddRf(A As VBProject, B As RfLin)
Dim F$: F = FrfeezRfLin(B)
If HasFrfee(A, F) Then Exit Sub
A.References.AddFromFile F
End Sub



Sub AddRfzAy(A As VBProject, RffAy$())
Dim F
For Each F In RffAy
    If Not HasFrfee(A, F) Then
        A.References.AddFromFile F
    End If
Next
End Sub
Sub CpyPjRfToPj(Pj As VBProject, ToPj As VBProject)
AddRfzAy ToPj, RffAyPj(Pj)
End Sub




Sub RmvPjRfNN(A As VBProject, RfNN$)
Dim N
For Each N In Ny(RfNN)
    'RmvPjRf A, N
Next
SavPj A
End Sub

Sub AddPjStdRf(A As VBProject, StdRfNm)
Const CSub$ = CMod & "AddPjStdRf"
If HasRfNm(A, StdRfNm) Then
    Debug.Print FmtQQ("AddPjStdRf: Pj(?) already has StdRfNm(?)", A.Name, StdRfNm)
    Exit Sub
End If
Dim Frfee$: Frfee = StdRff(StdRfNm)
ThwNotExistFfn Frfee, CSub, "StdRfFil"
A.References.AddFromFile Frfee
End Sub

