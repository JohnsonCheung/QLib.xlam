Attribute VB_Name = "QVb_Str_Apd"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Str_Apd."
Private Const Asm$ = "QVb"
Private Const Ns$ = "MVb_Str"

Function PpdSpcIf$(S)
PpdSpcIf = PpdIf(S, " ")
End Function

Function AddNBzAy(Ay, Sfx$) As String()
Dim I
For Each I In Itr(Ay)
    PushI AddNBzAy, AddNB(I, Sfx)
Next
End Function

Function PpdIf$(S, Pfx$)
If S = "" Then Exit Function
PpdIf = Pfx & S
End Function

Function AddNB$(ParamArray ApOf_Str())
'Ret : :S ! ret a str by adding each ele of @ApOf_Str one by one, if all them is <>'' else ret blank @@
Dim Av(): Av = ApOf_Str
Dim O$()
Dim S: For Each S In Itr(Av)
    If S = "" Then Exit Function
    Push O, S
Next
AddNB = Jn(O)
End Function
