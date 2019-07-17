Attribute VB_Name = "QVb_B_DocWs"
Option Explicit
Option Compare Text

Sub Z_DoTyDfnP()
BrwDrs DoTyDfnP
End Sub

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
    If Si(IGp) > 1 Then Stop
    PushI DyoTyDfn, DyoTyDfn__Dr(IGp)
Next
End Function
Private Function DyoTyDfn__Gp(Filter$()) As Variant()
Dim IsFstLin As Boolean, IsRstLin As Boolean
Dim Gp(), O()
Dim L: For Each L In Itr(Filter)
    DyoTyDfn__AsgWhatLin L, IsFstLin, IsRstLin
    Select Case True
    Case IsFstLin
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
OIsFstLin = DyoTyDfn__IsFstLin(RmkLin): If OIsFstLin Then Exit Sub
OIsRstLin = DyoTyDfn__IsRstLin(RmkLin)
End Sub
Private Function DyoTyDfn__IsFstLin(Lin) As Boolean
Dim T$(): T = T3Rst(Lin)
Dim NTerm%: NTerm = Si(T)
If NTerm < 2 Then Exit Function         ' At least 2 terms
If Fst2Chr(T(0)) <> "':" Then Exit Function '
If LasChr(T(0)) <> ":" Then Exit Function
If FstChr(T(1)) <> ":" Then Exit Function
Dim T3$: T3 = T(2)
Select Case FstChr(T3)
Case "!": DyoTyDfn__IsFstLin = True '<==
Case "#":
    If LasChr(T3) = "#" Then
        If NTerm >= 4 Then
            If FstChr(T(3)) = "!" Then
                DyoTyDfn__IsFstLin = True
            End If
        End If
    End If
End Select
End Function
Private Function DyoTyDfn__IsRstLin(Lin) As Boolean
If FstChr(Lin) <> "'" Then Exit Function
If FstChr(LTrim(RmvFstChr(Lin))) <> "!" Then Exit Function
DyoTyDfn__IsRstLin = True
End Function

Private Function DyoTyDfn__Dr(Gp) As Variant()
'Assume: Fst Lin is ':nn: :dd #mm# !rr
'        Rst Lin is '              !rr
Dim Nm$, Mem$, Dfn$, Rmk$, R$
Dim Fst As Boolean: Fst = True
Dim L: For Each L In Gp
    If Fst Then
        Fst = False
        Nm = Bef(Mid(L, 3), ":") ' Must
        R = AftSpc(L)
        Dfn = RmvFstChr(BefSpc(R)) 'Must
        R = AftSpc(R)
        If FstChr(R) = "#" Then
            Mem = RmvFstLasChr(T1(R))               'Optional
            R = RmvT1(R)
        End If
        Rmk = Aft(R, "!")
    Else
        ' L is in ' !rr fmt
        R = RmvFstChr(L)        ' Rmv '
        R = LTrim(R)
        R = RmvFstChr(R)        ' Rmv !
        Rmk = ApdIf(Rmk, vbCrLf & R)
    End If
Next
Nm = Qte(Nm, ":")
If Mem <> "" Then Mem = Qte(Mem, "#")
Dfn = ":" & Dfn
DyoTyDfn__Dr = Array(Nm, Dfn, Mem, Rmk)
End Function

