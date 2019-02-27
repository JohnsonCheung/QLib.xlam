Attribute VB_Name = "MVb_Dic_Rel"
Option Explicit

Function CvRel(A) As Rel
Set CvRel = A
End Function

Property Get EmpRel() As Rel
Set EmpRel = New Rel
End Property

Function IsRel(A) As Boolean
IsRel = TypeName(A) = "Rel"
End Function

Function Rel(RelLy$()) As Rel
Dim O As New Rel
Set Rel = O.Init(RelLy)
End Function

Function RelVbl(Vbl$) As Rel
Set RelVbl = Rel(SplitVBar(Vbl))
End Function
