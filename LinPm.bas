VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const CMod$ = "LinPm."
Public Dic As Dictionary ' it is a SyDic

Function Init(PmStr$) As LinPm
Dic.RemoveAll
Dim Ay$(): Ay = SySsl(PmStr)
    Dim I
    Dim CurPmNm$
    For Each I In Itr(Ay)
        If FstChr(I) = "-" Then
            CurPmNm = RmvFstChr(I)
            PushPmNm CurPmNm
        Else
            PushPm CurPmNm, CStr(I)
        End If
    Next
Set Init = Me
End Function
Private Sub PushPmNm(PmNm$)
If Dic.Exists(PmNm) Then Exit Sub
Dic.Add PmNm, Sy()
End Sub
Function WhNm(Optional NmPfx$) As WhNm
Set WhNm = MIde_Wh.WhNm(Patn(NmPfx), LikAy(NmPfx), ExlLikAy(NmPfx))
End Function
Function HasSw(SwNm) As Boolean
HasSw = HasEle(SwNy, SwNm)
End Function
Property Get SwNy() As String()
Dim K, V
For Each K In Dic.Keys
    V = Dic(K)
    If IsSy(Dic(K)) Then
        PushI SwNy, K
    End If
Next
End Property
Function SwNm$(Nm$)
If HasSw(Nm) Then SwNm = Nm
End Function
Property Get Cnt%()
Cnt = Dic.Count
End Property
Function HasPm(PmNm$) As Boolean
HasPm = Dic.Exists(PmNm)
End Function
Private Sub PushPm(Nm$, Optional V$)
Dim J%, S$()
If HasPm(Nm) Then
    If V = "" Then Exit Sub
    S = Dic(Nm)
    PushI S, V
    Dic(Nm) = S
Else
    Dic.Add Nm, Sy(V)
End If
End Sub
Sub Dmp()
D Fmt
End Sub
Function Fmt() As String()
Dim PmNm, O$()
For Each PmNm In Dic.Keys
    PushI O, FmtzNmSy(PmNm, CvSy(Dic(PmNm)))
Next
Fmt = FmtAyzSepSS(O, "ValCnt Val(")
End Function

Private Function FmtzNmSy$(PmNm, Sy$())
Select Case Si(Sy)
Case 0:    FmtzNmSy = FmtQQ("PmSw(?)", PmNm)
Case Else: FmtzNmSy = FmtQQ("Pm(?) ValCnt(?) Val(?)", PmNm, Si(Sy), JnSpc(Sy))
End Select
End Function

Function LikAy(NmPfx) As String()
LikAy = SyPmVal(NmPfx & "LikAy")
End Function

Function ExlLikAy(NmPfx) As String()
ExlLikAy = SyPmVal(NmPfx & "ExlLikAy")
End Function

Function SyPmVal(PmNm, Optional NmPfx$) As String()
If Dic.Exists(NmPfx & PmNm) Then
    SyPmVal = Dic(PmNm)
End If
End Function
Function StrPmVal$(PmNm, Optional NmPfx$)
Const CSub$ = CMod & "StrPmVal"
Dim Vy$()
    Vy = SyPmVal(PmNm, NmPfx)
Select Case Si(Vy)
Case 0
Case 1: StrPmVal = Vy(0)
Case Else: Thw CSub, FmtQQ("Parameter [-?] should have one value", PmNm), "Pm PmValSz PmVal-Sy", Fmt, Si(Vy), Vy
End Select
End Function
Function Patn$(NmPfx)
Patn = StrPmVal(NmPfx & "Patn")
End Function

Private Sub Class_Initialize()
Set Dic = New Dictionary
Dic.CompareMode = TextCompare
End Sub

