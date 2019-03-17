Attribute VB_Name = "Module2"
Option Explicit
Sub FmtLozStdWb(A As Workbook)
Dim Lo
For Each Lo In LoAy(A)
    FmtLozStd CvLo(Lo)
Next
End Sub
Sub FmtLozStd(A As ListObject)
FmtLo A, StdLof
End Sub
Property Get StdLof() As String()

End Property
