Attribute VB_Name = "MVb_Fs_Ft"
Option Explicit

Sub DmpFt(Ft$)
D LineszFt(Ft)
End Sub

Sub BrwFt(Ft$, Optional UseVc As Boolean)
Shell IIf(UseVc, "code.cmd", "notepad.exe") & " """ & Ft & """", vbMaximizedFocus
End Sub
Function LineszFt$(Ft$)
If SizFfn(Ft$) <= 0 Then Exit Function
LineszFt = Fso.GetFile(Ft$).OpenAsTextStream.ReadAll
End Function

Function LyzFt(Ft$) As String()
LyzFt = SplitCrLf(LineszFt(Ft))
End Function
Function FnoRnd128%(Ffn$)
FnoRnd128 = FnoRnd(Ffn, 128)
End Function

Function FnoRnd%(Ffn$, RecLen%)
Dim O%: O = FreeFile(1)
Open Ffn For Random As #O
FnoRnd = O
End Function

Function FnoApp%(Ft$)
Dim O%: O = FreeFile(1)
Open Ft For Append As #O
FnoApp = O
End Function

Function FnoInp%(Ft$)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FnoInp = O
End Function

Function FnoOup%(Ft$)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
FnoOup = O
End Function

Sub RmvFst4LinesFt(Ft$)
Dim A1$: A1 = Fso.GetFile(Ft$).OpenAsTextStream.ReadAll
Dim A2$: A2 = Left(A1, 55)
Dim A3$: A3 = Mid(A1, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If A2 <> B1 Then Stop
Fso.CreateTextFile(Ft$, True).Write A3
End Sub
