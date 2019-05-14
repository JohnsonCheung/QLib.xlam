Attribute VB_Name = "QVb_Run"
Option Explicit
Private Const CMod$ = "MVb_Run."
Private Const Asm$ = "QVb"
Function Pipe(Pm, Mthnn$)
Dim O: Asg Pm, O
Dim I
For Each I In Ny(Mthnn)
   Asg Run(I, O), O
Next
Asg O, Pipe
End Function

Function RunAvzIgnEr(Mthn, Av())
If Si(Av) > 9 Then Thw CSub, "Si(Av) should be 0-9", "Si(Av)", Si(Av)
On Error Resume Next
RunAv Mthn, Av
End Function
Function RunAv(Mthn, Av())
Dim O
Select Case Si(Av)
Case 0: O = Run(Mthn)
Case 1: O = Run(Mthn, Av(0))
Case 2: O = Run(Mthn, Av(0), Av(1))
Case 3: O = Run(Mthn, Av(0), Av(1), Av(2))
Case 4: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case 9: O = Run(Mthn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7), Av(8))
Case Else: Thw CSub, "UB-Av should be <= 8", "UB-Si Mthn", UB(Av), Mthn
End Select
RunAv = O
End Function

