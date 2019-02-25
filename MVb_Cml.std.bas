Attribute VB_Name = "MVb_Cml"
Option Explicit

Private Sub Z_CmlAset()
Debug.Print CmlAset(LinesPj(CurPj)).Cnt
End Sub
Function CmlAset(S) As Aset
Set CmlAset = AsetzAy(CmlAy(S))
End Function

Function Seg1ErNy() As String()
Erase XX
X "Act"
X "App"
X "Ass"
X "Ay"
X "Bar"
X "Brk"
X "C3"
X "C4"
X "Can"
X "Cell"
X "Cm"
X "Cmd"
X "Db"
X "Dbtt"
X "Dic"
X "Dry"
X "Ds"
X "Ent"
X "F"
X "Fb"
X "Fbq"
X "Fdr"
X "Fny"
X "Frm"
X "Fun"
X "Fx"
X "Git"
X "Has"
X "Lg"
X "Lgr"
X "Lnx"
X "Lo"
X "Md"
X "Min"
X "Msg"
X "Mth"
X "N"
X "O"
X "Pc"
X "Pj"
X "Ps1"
X "Pt"
X "Pth"
X "Re"
X "Res"
X "Rs"
X "Scl"
X "Sess"
X "Shp"
X "Spec"
X "Sql"
X "Sw"
X "T"
X "Tak"
X "Tim"
X "Tmp"
X "To"
X "Txtb"
X "V"
X "W"
X "Xls"
X "Y"
Seg1ErNy = XX
End Function


Function CmlxxAy(Ay) As String()
Dim L
For Each L In Itr(Ay)
    PushI CmlxxAy, Cmlxx(L)
Next
End Function

Function Cmlxx$(S)
Cmlxx = S & " " & JnSpc(CmlAy(S))
End Function
Private Sub Z_ShfCml()
Dim L$, EptL$
Ept = "A"
L = "AABcDD"
EptL = "ABcDD"
GoSub Tst
Exit Sub
Tst:
    Act = ShfCml(L)
    If Act <> Ept Then Stop
    If EptL <> L Then Stop
    Return
End Sub
Function FstCmlzWithSng$(S)
Dim Lin$, A$, O$, J%
Lin = S
While Lin <> ""
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    A = ShfCml(Lin)
    Select Case Len(A)
    Case 1: O = O & A
    Case Else: FstCmlzWithSng = O & A: Exit Function
    End Select
Wend
FstCmlzWithSng = O
End Function

Function FstCml$(S)
FstCml = ShfCml(CStr(S))
End Function

Function FstCmlx$(S)
FstCmlx = S & " " & FstCml(S)
End Function
Function FstCmlAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI FstCmlAy, FstCml(I)
Next
End Function
Function FstCmlxAy(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI FstCmlxAy, FstCmlx(I)
Next
End Function

Function ShfCml$(OStr$)
Dim J&, Fst As Boolean, Cml$, C$, A%
Fst = True
For J = 1 To Len(OStr)
    C = Mid(OStr, J, 1)
    A = Asc(C)
    Select Case True
    Case Fst
        Cml = C
        Fst = False
    Case IsAscUCase(A)
        If Cml <> "" Then GoTo R
        Cml = C
    Case IsAscDig(A)
        If Cml <> "" Then Cml = Cml & C
    Case IsAscLCase(A)
        Cml = Cml & C
    Case Else
        If Cml <> "" Then GoTo R
        Cml = ""
    End Select
Next
R:
    ShfCml = Cml
    OStr = Mid(OStr, J)
End Function

Function CmlAyShf(S) As String()
Dim L$: L = S
Dim J&
While True
    J = J + 1: If J > 100000 Then ThwLoopingTooMuch CSub
    PushNonBlankStr CmlAyShf, ShfCml(L)
    If L = "" Then Exit Function
Wend
End Function
Function CmlLy(Ny$()) As String()
Dim N
For Each N In Itr(Ny)
    PushI CmlLy, CmlLin(N)
Next
End Function
Function CmlLin$(Nm)
CmlLin = Nm & " " & JnSpc(CmlAy(Nm))
End Function
Function CmlAy(S) As String()
Dim J&, Fst As Boolean, Cml$, C$, A%
For J = 1 To Len(S)
    C = Mid(S, J, 1)
    A = Asc(C)
    Select Case True
    Case Fst
        Cml = C
        Fst = False
    Case IsAscUCase(A)
        PushNonBlankStr CmlAy, Cml
        Cml = C
    Case IsAscDig(A)
        If Cml <> "" Then Cml = Cml & C
    Case IsAscLCase(A)
        Cml = Cml & C
    Case Else
        PushNonBlankStr CmlAy, Cml
        Cml = ""
    End Select
Next
PushNonBlankStr CmlAy, Cml
End Function


