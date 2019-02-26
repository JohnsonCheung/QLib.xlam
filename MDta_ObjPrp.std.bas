Attribute VB_Name = "MDta_ObjPrp"
Option Explicit
Function DrszItrPP(Itr, PP_MayWith_NewFldEqQuoteFmFld$) As DRs
Dim A$(): A = SySsl(PP_MayWith_NewFldEqQuoteFmFld)
Dim PPzPrp$()
Dim PPzFml$()
Dim PPzAll$()
WAsg3PP PP_MayWith_NewFldEqQuoteFmFld, PPzPrp, PPzFml, PPzAll
ThwEr WFmlEr(PPzPrp, PPzFml), CSub
Dim Drs1 As DRs: Set Drs1 = DrszItrPPzPure(Itr, PPzPrp)
Dim Drs2 As DRs: Set Drs2 = WAddFml(Drs1, PPzFml)
Set DrszItrPP = DrsSel(Drs2, PPzAll)
End Function
Function DrszOyPP(Oy, PP_MayWith_NewFldEqQuoteFmFld$) As DRs
Set DrszOyPP = DrszItrPP(Itr(Oy), PP_MayWith_NewFldEqQuoteFmFld)
End Function

Private Function WFmlEr(PrpAy$(), PPzFml$()) As String()
Dim Fml, ErPmAy$(), PmAy$(), O$()
For Each Fml In Itr(PPzFml)
    PmAy = SplitComma(TakBetBkt(Fml))
    ErPmAy = AyMinus(PmAy, PrpAy)
    If Sz(ErPmAy) > 0 Then PushI O, FmtQQ("Invalid-Pm[?] in Fml[?]", JnSpc(ErPmAy), Fml)
Next
If Sz(O) > 0 Then PushI O, FmtQQ("Valid-Pm[?]", JnSpc(PrpAy))
WFmlEr = O
End Function
Private Sub WAsg3PP(PP_with_NewFldEqQuoteFmFld$, OPPzPrp$(), OPPzFml$(), OPPzAll$())
Dim I
For Each I In SySsl(PP_with_NewFldEqQuoteFmFld)
    If HasSubStr(I, "=") Then
        PushI OPPzAll, TakBef(I, "=")
        PushI OPPzFml, I
    Else
        PushI OPPzAll, I
        PushI OPPzPrp, I
    End If
Next
End Sub
Private Function WAddFml(A As DRs, PPzFml$()) As DRs
Dim O As DRs: Set O = A
Dim NewFld$, FunNm$, PmAy$(), Fml
For Each Fml In Itr(PPzFml)
    NewFld = TakBef(Fml, "=")
    FunNm = TakBet(Fml, "=", "(")
    PmAy = SplitComma(TakBetBkt(Fml))
    Set O = AddColzFmlDrs(O, NewFld, FunNm, PmAy)
Next
End Function
Function AddColzFmlDrs(A As DRs, NewFld, FunNm$, PmAy$()) As DRs
Dim Dry(): Dry = A.Dry
If Sz(Dry) = 0 Then Set AddColzFmlDrs = A: Exit Function
Dim Dr, U&, IxAy1&(), Av()
IxAy1 = IxAy(A.Fny, PmAy)
U = UB(A.Fny)
For Each Dr In Dry
    If UB(Dr) <> U Then Thw CSub, "Dr-Sz is diff", "Dr-Sz U", UB(Dr), U
    Av = AywIxAy(Dr, IxAy1)
    Push Dr, RunAv(FunNm, Av)
Next
Set AddColzFmlDrs = DRs(AyAddItm(A.Fny, NewFld), Dry)
End Function
Private Function DrszItrPPzPure(Oy, PP) As DRs
Set DrszItrPPzPure = DRs(PP, DryzItrPPzPure(Oy, PP))
End Function

Private Function DryzItrPPzPure(Itr, PP) As Variant()
Dim U%, I
Dim PrpNy$()
PrpNy = CvNy(PP)
For Each I In Itr
    Push DryzItrPPzPure, DrzObjPrpNy(I, PrpNy)
Next
End Function

Private Sub Z_DrszItrPP()
'BrwDrs DrszItrPP(Excel.Application.AddIns, "Name Installed IsOpen FullName CLSId ")
'BrwDrs DrszItrPP(Fds(Db(SampFbzDutyDta), "Permit"), "Name Type Required")
'BrwDrs ItrPrpDrs(Application.VBE.VBProjects, "Name Type")
BrwDrs DrszItrPP(CurPj.VBComponents, "Name Type CmpTy=ShpCmpTy(Type)")
End Sub

Private Sub Z()
Z_DrszItrPP
MDta_ObjPrp:
End Sub

