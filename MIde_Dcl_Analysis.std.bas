Attribute VB_Name = "MIde_Dcl_Analysis"
Option Explicit

Function DclT1$(A)
Dim L$: L = LTrim(A)
If L = "" Then Exit Function
If FstChr(L) = "'" Then Exit Function
DclT1 = T1(A)
End Function
Function DclLy_T1Aset(A$()) As Aset
Dim L, O As Aset
Set O = EmpAset
For Each L In Itr(A)
'    AsetPush O, DclT1(L)
Next
End Function
Function Md_DclLinT1Ay(A As CodeModule) As String()
Md_DclLinT1Ay = DclLy_T1Aset(DclLyMd(A))
End Function
