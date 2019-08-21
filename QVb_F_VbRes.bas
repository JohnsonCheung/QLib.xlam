Attribute VB_Name = "QVb_F_VbRes"
Option Explicit
Option Compare Text
Private X_ResPseg$
Property Get ResPseg()
ResPseg = X_ResPseg
End Property
Function IsPseg(Pseg$)
Select Case True
Case FstChr(Pseg) = "\"
Case LasChr(Pseg) = "\"
Case Else: IsPseg = True
End Select
End Function

Sub AssIsPseg(Pseg$, Optional Fun$ = "AssIsPseg")
If IsPseg(Pseg) Then Exit Sub
Thw Fun, "Given Pseg not not Pth-seg", "Pseg", Pseg
End Sub

Sub SetResPseg(ResPseg$)
If ResPseg = "" Then Exit Sub
If X_ResPseg = ResPseg Then Exit Sub
AssIsPseg ResPseg
Dim O$: O = ResHom & ResPseg
X_ResPseg = ResPseg
EnsPthzAllSeg O
End Sub

Function Res$(ResFn$, Optional ResPseg$)
Res = JnCrLf(ResLy(ResFn, ResPseg))
End Function

Function WrtRes$(S$, ResFn$, Optional OvrWrt As Boolean)
Dim Ft$: Ft = FtzRes(ResFn)
WrtStr S, Ft, OvrWrt
WrtRes = Ft
End Function

Function ResLy(ResFn$, Optional ResPseg$) As String()
ResLy = LyzFt(FtzRes(ResFn, ResPseg))
End Function

Function ResHom$()
ResHom = AddFdrEns(TmpHom, "Res")
End Function

Function FtzRes$(ResFn$, Optional ResPseg$)
'Ret : :<ResPseg><PthSep><ResFn> #Ft-Of-Res# @@
':Pseg: :S #Pth-Segment# ! zero of more DirSeg sep by pthSep
':DirSeg: :S ! EntNm added in OS file system. '.' & '..' are allowed.
SetResPseg ResPseg
Dim O$: O = ResHom & ResPseg & PthSep & ResFn
EnsPthzAllSeg O
FtzRes = O
End Function

