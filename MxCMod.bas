Attribute VB_Name = "MxCMod"
Option Explicit
Option Compare Text
Const CNs$ = "AA"
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxCMod."
Public Const FFoCMod$ = "CLibv CNsv CModv"

':CNsv: :S #Cnst-CNs-Value# ! the string bet-DblQ of CnstLin-CNs of a Md
':CModv: :S #Cnst-CMod-Value# ! the string aft-rmv-sfx-[.] of bet-DblQ of CnstLin-CMod of a Md
':CLibv: :S #Cnst-CLib-Value# ! the string aft-rmv-sfx-[.] of bet-DblQ of CnstLin-CLib of a Md
Function FoCMod() As String()
FoCMod = SyzSS(FFoCMod)
End Function

Function CNsv$(M As CodeModule)
CNsv = BetDblQ(CnstLLin(M, "CNs").Lin)
End Function

Sub RmvCModP()
RmvCModzP CPj
End Sub

Sub RmvCModzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    RmvCModzM C.CodeModule
Next
End Sub

Function DroCMod(M As CodeModule)
DroCMod = Array(CLibv(M), CNsv(M), CModv(M))
End Function

Function CModv$(M As CodeModule)
CModv = RmvSfxDot(BetDblQ(CnstLLin(M, "CMod").Lin))
End Function

Function CLibv$(M As CodeModule)
CLibv = RmvSfxDot(BetDblQ(CnstLLin(M, "CLib").Lin))
End Function

Sub ClrCModM()
ClrCModzM CMd
End Sub

Sub ClrCModzM(M As CodeModule)
ClrCnstLin M, "CMod"
End Sub

Sub ClrCLibzM(M As CodeModule)
ClrCnstLin M, "CMod"
End Sub

Sub RmvCModM()
RmvCModzM CMd
End Sub

Sub RmvCModzM(M As CodeModule)
RmvCnstLin M, "CMod"
End Sub

Sub RmvCLibM()
RmvCLibzM CMd
End Sub

Sub RmvCLibzM(M As CodeModule)
RmvCnstLin M, "CLib"
End Sub

Sub RmvCLibP()
RmvCLibzP CPj
End Sub

Sub RmvCLibzP(P As VBProject)
RmvCnstLinzP P, "CLib", IsPrvOnly:=True
End Sub

Sub EntCLibP()
EntCLibzP CPj
End Sub
Sub EntCLibzP(P As VBProject)
Dim A As New InpVal
Dim C As VBComponent: For Each C In P.VBComponents
    EntCLibzM C.CodeModule, A
Next
End Sub

Sub EntCLibM()
Dim A As New InpVal
A.Ini Tit:="Md[" & CMdn & "]", Prompt:="Enter CLib value"
EntCLibzM CMd, A
End Sub

Sub EntCLibzM(M As CodeModule, A As ValInp)
If Not IsMd(M.Parent) Then Exit Sub
Dim V$
V = CLibv(M): If V <> "" Then Exit Sub
V = A.Inp
If FstChr(V) <> "Q" Then Exit Sub
EnsCnstLin M, CLibLin(V)
End Sub

Sub EnsCLibzM(M As CodeModule, CLibv$)
If Not IsMd(M.Parent) Then Exit Sub
EnsCnstLin M, CLibLin(CLibv)
End Sub

Sub EnsCNsLin(M As CodeModule, Ns$)
If Not IsMd(M.Parent) Then Exit Sub
EnsCnstLin M, CNsLin(Ns)
End Sub

Sub EnsCNszM(M As CodeModule, Ns$)
EnsCnstLin M, CNsLin(Ns)
End Sub

Function CNsLin$(Ns$)
':CLibLin: :PrvCnstLin ! Is a `Const CLib$ = "${Clibv}."`
CNsLin = FmtQQ("Const CNs$ = ""?""", Ns)
End Function

Function CLibLin$(CLibv$)
':CLibLin: :PrvCnstLin ! Is a `Const CLib$ = "${Clibv}."`
CLibLin = FmtQQ("Const CLib$ = ""?.""", CLibv)
End Function

Sub EnsCModM()
EnsCModzM CMd
End Sub

Sub EnsCModP()
EnsCModzP CPj
End Sub

Sub EnsCModzM(M As CodeModule)
EnsCnstLinAft M, CModLin(M), "CLib", IsPrvOnly:=True
End Sub


Function CModLin$(M As CodeModule)
':CModLin: :CnstLin ! Is a Const CMod$ = CLib & "xxxx."
CModLin = FmtQQ("Const CMod$ = CLib & ""?.""", Mdn(M))
End Function

Sub EnsCModzP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    EnsCModzM C.CodeModule
Next
End Sub
