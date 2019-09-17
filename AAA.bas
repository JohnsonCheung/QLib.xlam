Attribute VB_Name = "AAA"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "AAA."

Function DymJnDot(Dy()) As String()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DymJnDot, JnDot(Dr)
Next
End Function
