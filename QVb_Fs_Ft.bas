Attribute VB_Name = "QVb_Fs_Ft"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Fs_Ft."
Private Const Asm$ = "QVb"

Sub DmpFt(Ft)
D LineszFt(Ft)
End Sub
Function EnsFt$(Ft$)
If Not HasFfn(Ft) Then
    WrtStr "", Ft
End If
EnsFt = Ft
End Function
Sub BrwFt(Ft, Optional UseVc As Boolean)
Shell IIf(UseVc, "code.cmd", "notepad.exe") & " """ & Ft & """", vbMaximizedFocus
End Sub
Function LineszFt$(Ft)
LineszFt = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
End Function
Sub CrtFfn(Ffn)
Close #FnoO(Ffn)
End Sub

Function LyzFt(Ft) As String()
LyzFt = SplitCrLf(LineszFt(Ft))
End Function
Function FnoRnd128%(Ffn)
FnoRnd128 = FnoRnd(Ffn, 128)
End Function

Function FnoRnd%(Ffn, RecLen%)
Dim O%: O = FreeFile(1)
Open Ffn For Random As #O
FnoRnd = O
End Function

Function FnoA%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Append As #O
FnoA = O
End Function

Function FnoI%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FnoI = O
End Function

Function FnoO%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
FnoO = O
End Function

Sub RmvFst4LinesFt(Ft)
Dim A1$: A1 = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
Dim A2$: A2 = Left(A1, 55)
Dim A3$: A3 = Mid(A1, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If A2 <> B1 Then Stop
Fso.CreateTextFile(Ft, True).Write A3
End Sub
