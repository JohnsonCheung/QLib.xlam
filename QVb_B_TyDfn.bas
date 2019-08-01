Attribute VB_Name = "QVb_B_TyDfn"
Option Explicit
Option Compare Text

Private Sub Z_DoTyDfnP()
BrwDrs DoTyDfnP
End Sub

Function TyDfnNyP() As String()
TyDfnNyP = TyDfnNyzP(CPj)
End Function

Function TyDfnNyzP(P As VBProject) As String()
Dim L: For Each L In RmkLy(SrczP(P))
    PushNB TyDfnNyzP, TyDfnNm(L)
Next
End Function

Function TyDfnNm$(L)
Dim T$: T = T1(L)
If T = "" Then Exit Function
If Fst2Chr(T) <> "':" Then Exit Function
If LasChr(T) <> ":" Then Exit Function
TyDfnNm = RmvFstChr(T)
End Function

Function DoTyDfnP() As Drs
':DoTyDfn: :Drs<Nm-Ty-Mem-Rmk>
DoTyDfnP = DoTyDfnzP(CPj)
End Function

Private Function DoTyDfnzP(P As VBProject) As Drs
DoTyDfnzP = Drs(FoTyDfn, DyoTyDfn(RmkLy(SrczP(P))))
End Function

Private Function FoTyDfn() As String()
FoTyDfn = SyzSS("Nm Ty Mem Rmk")
End Function

Function WsoTyDfn() As Worksheet
Set WsoTyDfn = WsoTyDfnzP(CPj)
End Function

Function WsoTyDfnzP(P As VBProject) As Worksheet
Dim O As New Worksheet
Set O = NewWs("TyDfn")
'RgzSq DocSqzP(P), A1zWs(O)
Stop
FmtWsoTyDfn O
Set WsoTyDfnzP = O
End Function

Private Sub FmtWsoTyDfn(WsoTyDfn As Worksheet)

End Sub

Private Function SqoTyDfnzP(P As VBProject) As Variant()
SqoTyDfnzP = SqzDy(DyoTyDfn(RmkLy(SrczP(P))))
End Function

Private Function DyoTyDfn(RmkLy$()) As Variant()
':DyoTyDfn: :Dyo<Nm-Ty-Mem-Rmk> #Dyo-TyDfn# ! Fst-Lin must be :nn: :dd #mm# !rr
'                                                           ! Rst-Lin is !rr
'                                                           ! must term: nn dd mm, all of them has no spc
'                                                           ! opt      : mm rr
'                                                           ! :xx:     : should uniq in pj
Dim Gp(): Gp = DyoTyDfn__Gp(RmkLy)
Dim I%:
Dim IGp: For Each IGp In Itr(Gp)
    PushSomSi DyoTyDfn, DyoTyDfn__Dr(IGp)
Next
End Function
Private Function DyoTyDfn__Gp(RmkLy$()) As Variant()
Dim IsFstLin As Boolean, IsRstLin As Boolean
Dim Gp(), O()
Dim L: For Each L In Itr(RmkLy)
    Dim NFstLin%
    DyoTyDfn__AsgWhatLin L, IsFstLin, IsRstLin
    Select Case True
    Case IsFstLin
        NFstLin = NFstLin + 1
        PushSomSi O, Gp
        Erase Gp
        PushI Gp, L
    Case IsRstLin
        If Si(Gp) > 0 Then
            PushI Gp, L ' Only with Fst-Lin, the Rst-Lin will be use, otherwise ign it.
        End If
    Case Else
        PushSomSi O, Gp
        Erase Gp
    End Select
Next
DyoTyDfn__Gp = O
End Function

Private Sub DyoTyDfn__AsgWhatLin(RmkLin, OIsFstLin As Boolean, OIsRstLin As Boolean)
OIsRstLin = False
OIsRstLin = False
OIsFstLin = TyDfnNm(RmkLin) <> "": If OIsFstLin Then Exit Sub
OIsRstLin = DyoTyDfn__IsRstLin(RmkLin)
End Sub

Private Function DyoTyDfn__IsRstLin(Lin) As Boolean
If FstChr(Lin) <> "'" Then Exit Function
If FstChr(LTrim(RmvFstChr(Lin))) <> "!" Then Exit Function
DyoTyDfn__IsRstLin = True
End Function

Private Sub Z_DyoTyDfn__Dr()
Dim RmkLy
GoSub ZZ
Exit Sub
ZZ:
    RmkLy = Sy("':Cell: :SCell-or-:WCell")
    Dmp DyoTyDfn__Dr(RmkLy)
    Return
End Sub

Private Function IsTyDfn(Nm$, Dfn$, T3$, Rst$) As Boolean
Select Case True
Case Fst2Chr(Nm) <> "':"
Case LasChr(Nm) <> ":"
Case FstChr(Dfn) <> ":"
Case T3 <> "" And Not HasPfxSfx(T3, "#", "#") And FstChr(T3) <> "!"
Case Else: IsTyDfn = True
End Select
End Function

Private Function IsLinTyDfn(Lin) As Boolean
Dim Nm$, Dfn$, T3$, Rst$
Asg3TRst Lin, Nm, Dfn, T3, Rst
If Not IsTyDfn(Nm, Dfn, T3, Rst) Then Exit Function
End Function

Private Function DyoTyDfn__Dr(RmkLy) As Variant()
'Assume: Fst Lin is ':nn: :dd [#mm#] [!rr]
'        Rst Lin is '                 !rr
Dim Nm$, T3$, Dfn$, Mem$, Rst$, Rmk$, R$
Dim Fst As Boolean: Fst = True
Dim L: For Each L In RmkLy
    If Fst Then
        Fst = False
        Asg3TRst L, Nm, Dfn, T3, Rst
        If Not IsTyDfn(Nm, Dfn, T3, Rst) Then
            Debug.Print Nm, Dfn, T3, Rst
            Stop
            Debug.Print "DyoTyDfn__Dr: Fst lin of @RmkLy is not ':nn: :dd [#mm#] [!rr].  FstLin=[" & L & "]"
            Exit Function
        End If
        If HasPfxSfx(T3, "#", "#") Then Mem = T3
        Rmk = Trim(RmvPfx(Rst, "!"))
    Else
        ' L is in ' !rr fmt
        R = RmvFstChr(L)        ' Rmv '
        R = LTrim(R)
        R = RmvFstChr(R)        ' Rmv !
        Rmk = ApdIf(Rmk, vbCrLf & R)
    End If
Next
Nm = RmvFstChr(Nm)
DyoTyDfn__Dr = Array(Nm, Dfn, Mem, Rmk)
End Function
