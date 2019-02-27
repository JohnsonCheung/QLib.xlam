Attribute VB_Name = "MIde_Gen_Rf_Add"
Option Explicit
Const CMod$ = "MIde_Pj_Rf_Add."

Private Sub ZZ()
Dim A As VBProject
Dim B As Variant
Dim C$()
Dim D$
Dim E&
Dim G As Dictionary
AddRfzAy A, C
End Sub

Private Sub Z()
End Sub
Sub AddRfzRff(A As VBProject, Rff)
Const CSub$ = CMod & "AddRf"
If HasRff(A, Rff) Then
    InfoLin CSub, "Rff exists in Pj", "Rff Pj", Rff, A.Name
    Exit Sub
End If
A.References.AddFromFile Rff
InfoLin CSub, "Rff is added to Pj", "Rff Pj", Rff, A.Name
End Sub

Sub AddRfzAy(A As VBProject, RffAy$())
Dim F
For Each F In RffAy
    If Not HasRff(A, F) Then
        A.References.AddFromFile F
    End If
Next
End Sub
