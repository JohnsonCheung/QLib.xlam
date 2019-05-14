Attribute VB_Name = "QVb_Dta_Lpm"
Type Lpm
    DicK_ToSy As Dictionary 'Its name is LpmDic
End Type
Function Lpm(PmStr$, LpmSpec$) As Lpm
Dim O As New Dictionary
Lpm.DicK_ToSy = O
Dim Ay$(): Ay = SyzSS(PmStr)
    Dim I, S$
    Dim CurPmNm$
    For Each I In Itr(Ay)
        S = I
        If FstChr(S) = "-" Then
            CurPmNm = RmvFstChr(S)
            PushPmNm CurPmNm
        Else
            PushPm CurPmNm, S
        End If
    Next
Set Init = Me
ThwIf_LpmEr O, LpmSpec
End Function
Private Sub ThwIf_LpmEr(LpmDic As Dictionary, LpmSpec$)

End Sub
Private Sub PushPmNm(LpmDic As Dictionary, PmNm$)
If Dic.Exists(PmNm) Then Exit Sub
Dic.Add PmNm, Sy()
End Sub

Function WhNm(Optional NmPfx$) As WhNm
Set WhNm = QIde_Wh.WhNm(Patn(NmPfx), LikeAy(NmPfx), ExlLikSy(NmPfx))
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
Function SwNm$(Nm)
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
Fmt = AlignzBySepss(O, "ValCnt Val(")
End Function

Private Function FmtzNmSy$(PmNm, Sy$())
Select Case Si(Sy)
Case 0:    FmtzNmSy = FmtQQ("PmSw(?)", PmNm)
Case Else: FmtzNmSy = FmtQQ("Pm(?) ValCnt(?) Val(?)", PmNm, Si(Sy), JnSpc(Sy))
End Select
End Function

Function LikeAy(NmPfx) As String()
LikeAy = SyPmVal(NmPfx & "LikeAy")
End Function

Function ExlLikSy(NmPfx) As String()
ExlLikSy = SyPmVal(NmPfx & "ExlLikSy")
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
Case Else: Thw CSub, FmtQQ("Parameter [-?] should have one value", PmNm), "Pm PmValSz ValzPm-Sy", Fmt, Si(Vy), Vy
End Select
End Function
Function Patn$(NmPfx)
Patn = StrPmVal(NmPfx & "Patn")
End Function

Private Sub Class_Initialize()
Set Dic = New Dictionary
Dic.CompareMode = TextCompare
End Sub


