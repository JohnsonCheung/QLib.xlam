Attribute VB_Name = "MXls_Lo_Set"
Option Explicit
Function WszLo(A As ListObject) As Worksheet
Set WszLo = A.Parent
End Function

Sub SetLoNm(A As ListObject, LoNm$)
If LoNm <> "" Then
    If Not HasLo(WszLo(A), LoNm) Then
        A.Name = LoNm
    Else
        Inf CSub, "Lo"
    End If
End If
End Sub


Sub SetLccWrp(A As ListObject, CC$, Optional Wrp As Boolean)
Dim C
For Each C In NyzNN(CC)
    SetLcWrp A, C, Wrp
Next
End Sub
Sub SetLcWrp(A As ListObject, C, Optional Wrp As Boolean)
A.ListColumns(C).DataBodyRange.WrapText = Wrp
End Sub

Sub SetLccWdt(A As ListObject, CC$, W)
Dim C
For Each C In NyzNN(CC)
    SetLcWdt A, C, W
Next
End Sub
Sub SetLcWdt(A As ListObject, C, W)
EntColzLc(A, C).ColumnWidth = W
End Sub
Function EntColzLc(A As ListObject, C) As Range
Set EntColzLc = A.ListColumns(C).DataBodyRange.EntireColumn
End Function
Sub SetLcTotLnk(A As ListObject, C)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = A.ListColumns(C).DataBodyRange
Set Ws = WszRg(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, R.Rows.Count + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

