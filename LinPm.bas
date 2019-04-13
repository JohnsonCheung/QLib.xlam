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

Private Sub ThwIfPmStrEr(PmNmToPmValSyDic As Dictionary, LinPmSpec$)

End Sub

Function Init(PmStr$, LinPmSpec$) As LinPm
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
ThwIfPmStrEr Dic, LinPmSpec
End Function
Private Sub PushPmNm(Pmnm$)
If Dic.Exists(Pmnm) Then Exit Sub
Dic.Add Pmnm, Sy()
End Sub
Function WhNm(Optional NmPfx$) As WhNm
Set WhNm = MIde_Wh.WhNm(Patn(NmPfx), LikeAy(NmPfx), ExlLikAy(NmPfx))
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
Function HasPm(Pmnm$) As Boolean
HasPm = Dic.Exists(Pmnm)
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
Dim Pmnm, O$()
For Each Pmnm In Dic.Keys
    PushI O, FmtzNmSy(Pmnm, CvSy(Dic(Pmnm)))
Next
Fmt = FmtAyzSepSS(O, "ValCnt Val(")
End Function

Private Function FmtzNmSy$(Pmnm, Sy$())
Select Case Si(Sy)
Case 0:    FmtzNmSy = FmtQQ("PmSw(?)", Pmnm)
Case Else: FmtzNmSy = FmtQQ("Pm(?) ValCnt(?) Val(?)", Pmnm, Si(Sy), JnSpc(Sy))
End Select
End Function

Function LikeAy(NmPfx) As String()
LikeAy = SyPmVal(NmPfx & "LikeAy")
End Function

Function ExlLikAy(NmPfx) As String()
ExlLikAy = SyPmVal(NmPfx & "ExlLikAy")
End Function

Function SyPmVal(Pmnm, Optional NmPfx$) As String()
If Dic.Exists(NmPfx & Pmnm) Then
    SyPmVal = Dic(Pmnm)
End If
End Function
Function StrPmVal$(Pmnm, Optional NmPfx$)
Const CSub$ = CMod & "StrPmVal"
Dim Vy$()
    Vy = SyPmVal(Pmnm, NmPfx)
Select Case Si(Vy)
Case 0
Case 1: StrPmVal = Vy(0)
Case Else: Thw CSub, FmtQQ("Parameter [-?] should have one value", Pmnm), "Pm PmValSz PmVal-Sy", Fmt, Si(Vy), Vy
End Select
End Function
Function Patn$(NmPfx)
Patn = StrPmVal(NmPfx & "Patn")
End Function

Private Sub Class_Initialize()
Set Dic = New Dictionary
Dic.CompareMode = TextCompare
End Sub

