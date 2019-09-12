Attribute VB_Name = "MxRfOp"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxRfOp."
Sub AddRfzRff(P As VBProject, Frfee$)
Const CSub$ = CMod & "AddRf"
If HasFrfee(P, Frfee) Then
    InfLin CSub, "Frfee exists in Pj", "Frfee Pj", Frfee, P.Name
    Exit Sub
End If
P.References.AddFromFile Frfee
InfLin CSub, "Frfee is added to Pj", "Frfee Pj", Frfee, P.Name
End Sub

Sub AddRfzS(P As VBProject, RfSrc$())
Dim I
For Each I In Itr(RfSrc)
    AddRf P, RfLin(CStr(I))
Next
End Sub

Sub AddRf(P As VBProject, B As RfLin)
Dim F$: F = FrfeezRfLin(B)
If HasFrfee(P, F) Then Exit Sub
P.References.AddFromFile F
End Sub



Sub AddRfzAy(P As VBProject, RffAy$())
Dim F
For Each F In RffAy
    If Not HasFrfee(P, F) Then
        P.References.AddFromFile F
    End If
Next
End Sub
Sub CpyPjRfToPj(P As VBProject, ToPj As VBProject)
AddRfzAy ToPj, RffyzP(P)
End Sub
Sub RmvRf(P As VBProject, Rfn)

End Sub
Sub RmvRfzPRr(P As VBProject, RfNN$)
Dim N
For Each N In Ny(RfNN)
    RmvRf P, N
Next
SavPj P
End Sub

Sub AddPjStdRf(P As VBProject, StdRfNm)
Const CSub$ = CMod & "AddPjStdRf"
'If HasRfNm(P, StdRfNm) Then
    Debug.Print FmtQQ("AddPjStdRf: Pj(?) already has StdRfNm(?)", P.Name, StdRfNm)
    Exit Sub
'End If
Dim Frfee$: 'Frfee = StdRff(StdRfNm)
'ThwNotExistFfn Frfee, CSub, "StdRfFil"
P.References.AddFromFile Frfee
End Sub
