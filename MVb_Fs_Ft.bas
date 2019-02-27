Attribute VB_Name = "MVb_Fs_Ft"
Option Explicit
Sub BrwFtVC(Ft)
Shell "code.cmd """ & Ft & """", vbHide
End Sub

Sub BrwFt(Ft)
Shell "notepad.exe """ & Ft & """", vbMaximizedFocus
End Sub
Function FtLines$(A)
If FfnSz(A) <= 0 Then Exit Function
FtLines = Fso.GetFile(A).OpenAsTextStream.ReadAll
End Function

Function FtLy(A) As String()
FtLy = SplitCrLf(FtLines(A))
End Function
Function FnoRnd128%(Ffn)
FnoRnd128 = FnoRnd(Ffn, 128)
End Function

Function FnoRnd%(Ffn, RecLen%)
Dim O%: O = FreeFile(1)
Open Ffn For Random As #O
FnoRnd = O
End Function

Function FnoApp%(A)
Dim O%: O = FreeFile(1)
Open A For Append As #O
FnoApp = O
End Function

Function FnoInp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FnoInp = O
End Function

Function FnoOup%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
FnoOup = O
End Function

Sub RmvFst4LinesFt(Ft)
Dim A1$: A1 = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
Dim A2$: A2 = Left(A1, 55)
Dim A3$: A3 = Mid(A1, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If A2 <> B1 Then Stop
Fso.CreateTextFile(Ft, True).Write A3
End Sub
