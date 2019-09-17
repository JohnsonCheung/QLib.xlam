Attribute VB_Name = "MxRetItrItm"
Option Explicit
Option Compare Text
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxRetItrItm."
Function ItwNm(Itr, Nm)
Dim O: For Each O In Itr
    If O.Name = Nm Then Asg O, ItwNm
Next
End Function
