Attribute VB_Name = "QDta_ObjPrp"
Option Explicit
Private Const CMod$ = "MDta_ObjPrp."
Private Const Asm$ = "QDta"

Function DrszItrPrpPthSy(Itr, PrpPthSy$()) As Drs
DrszItrPrpPthSy = Drs(PrpPthSy, DryzItrPrpPthSy(Itr, PrpPthSy))
End Function

Private Function WFmlEr(PrpVy$(), PPzFml$()) As String()
Dim Fml, ErPmAy$(), PmAy$(), O$()
For Each Fml In Itr(PPzFml)
    PmAy = SplitComma(BetBkt(Fml))
    ErPmAy = MinusAy(PmAy, PrpVy)
    If Si(ErPmAy) > 0 Then PushI O, FmtQQ("Invalid-Pm[?] in Fml[?]", JnSpc(ErPmAy), Fml)
Next
If Si(O) > 0 Then PushI O, FmtQQ("Valid-Pm[?]", JnSpc(PrpVy))
WFmlEr = O
End Function

Private Sub WAsg3PP(PP_with_NewFldEqQuoteFmFld$, OPPzPrp$(), OPPzFml$(), OPPzAll$())
Dim I, S$
For Each I In SyzSsLin(PP_with_NewFldEqQuoteFmFld)
    S = I
    If HasSubStr(S, "=") Then
        PushI OPPzAll, Bef(S, "=")
        PushI OPPzFml, I
    Else
        PushI OPPzAll, I
        PushI OPPzPrp, I
    End If
Next
End Sub

Private Function AddFmlSy(A As Drs, FmlSy$()) As Drs
Dim O As Drs: O = A
Dim NewFld$, FunNm$, PmAy$(), Fml$, I
For Each I In Itr(FmlSy)
    Fml = I
    NewFld = Bef(Fml, "=")
    FunNm = Bet(Fml, "=", "(")
    PmAy = SplitComma(BetBkt(Fml))
    O = AddFml(O, NewFld, FunNm, PmAy)
Next
End Function

Function AddFml(A As Drs, NewFld$, FunNm$, PmAy$()) As Drs
Dim Dry(): Dry = A.Dry
If Si(Dry) = 0 Then AddFml = A: Exit Function
Dim Dr, U&, IxAy1&(), Av()
IxAy1 = IxAy(A.Fny, PmAy)
U = UB(A.Fny)
For Each Dr In Dry
    If UB(Dr) <> U Then Thw CSub, "Dr-Si is diff", "Dr-Si U", UB(Dr), U
    Av = AywIxAy(Dr, IxAy1)
    Push Dr, RunAv(FunNm, Av)
Next
AddFml = Drs(AddSyItm(A.Fny, NewFld), Dry)
End Function

Function DryzItrPrpPthSy(Itr, PrpPthSy$()) As Variant()
Dim Obj As Object
For Each Obj In Itr
    Push DryzItrPrpPthSy, DrzObjPrpPthSy(Obj, PrpPthSy)
Next
End Function

Private Sub Z_DrszItrPrpPthSy()
'BrwDrs DrszItrPP(Excel.Application.Addins, "Name Installed IsOpen FullName CLSId ")
'BrwDrs DrszItrPP(Fds(Db(SampFbzDutyDta), "Permit"), "Name Type Required")
'BrwDrs ItrPrpDrs(Application.VBE.VBProjects, "Name Type")
BrwDrs DrszItrPrpPthSy(CurPj.VBComponents, SyzSsLin("Name Type CmpTy=ShpCmpTy(Type)"))
End Sub

Private Sub Z()
MDta_Prp:
End Sub

