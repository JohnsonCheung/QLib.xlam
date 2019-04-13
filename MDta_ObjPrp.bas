Attribute VB_Name = "MDta_ObjPrp"
Option Explicit
Sub AX()
Dim B
Set B = CurPj
Stop
End Sub
Function DrszItr(Itr, PP$) As Drs
Dim PrpPthAy$(): PrpPthAy = Ny(PP)
Set DrszItr = Drs(PrpPthAy, DryzItr(Itr, PrpPthAy))
End Function

Function DrszItrpp(Itr, PP) As Drs
Dim A$(): A = NyzNN(PP)
Set DrszItrpp = Drs(A, DryzItr(Itr, A))
End Function

Function DrszOypp(Oy, PP) As Drs
Set DrszOypp = DrszItrpp(Itr(Oy), PP)
End Function

Private Function WFmlEr(PrpAy$(), PPzFml$()) As String()
Dim Fml, ErPmAy$(), PmAy$(), O$()
For Each Fml In Itr(PPzFml)
    PmAy = SplitComma(BetBkt(Fml))
    ErPmAy = AyMinus(PmAy, PrpAy)
    If Si(ErPmAy) > 0 Then PushI O, FmtQQ("Invalid-Pm[?] in Fml[?]", JnSpc(ErPmAy), Fml)
Next
If Si(O) > 0 Then PushI O, FmtQQ("Valid-Pm[?]", JnSpc(PrpAy))
WFmlEr = O
End Function

Private Sub WAsg3PP(PP_with_NewFldEqQuoteFmFld$, OPPzPrp$(), OPPzFml$(), OPPzAll$())
Dim I
For Each I In SySsl(PP_with_NewFldEqQuoteFmFld)
    If HasSubStr(I, "=") Then
        PushI OPPzAll, StrBef(I, "=")
        PushI OPPzFml, I
    Else
        PushI OPPzAll, I
        PushI OPPzPrp, I
    End If
Next
End Sub

Private Function DrsAddFml(A As Drs, PPzFml$()) As Drs
Dim O As Drs: Set O = A
Dim NewFld$, FunNm$, PmAy$(), Fml
For Each Fml In Itr(PPzFml)
    NewFld = StrBef(Fml, "=")
    FunNm = StrBet(Fml, "=", "(")
    PmAy = SplitComma(BetBkt(Fml))
    Set O = AddColzFmlDrs(O, NewFld, FunNm, PmAy)
Next
End Function

Function AddColzFmlDrs(A As Drs, NewFld, FunNm$, PmAy$()) As Drs
Dim Dry(): Dry = A.Dry
If Si(Dry) = 0 Then Set AddColzFmlDrs = A: Exit Function
Dim Dr, U&, IxAy1&(), Av()
IxAy1 = IxAy(A.Fny, PmAy)
U = UB(A.Fny)
For Each Dr In Dry
    If UB(Dr) <> U Then Thw CSub, "Dr-Si is diff", "Dr-Si U", UB(Dr), U
    Av = AywIxAy(Dr, IxAy1)
    Push Dr, RunAv(FunNm, Av)
Next
Set AddColzFmlDrs = Drs(AyAddItm(A.Fny, NewFld), Dry)
End Function

Private Function DryzItr(Itr, PrpPthAy$()) As Variant()
Dim Obj
For Each Obj In Itr
    Push DryzItr, DrzObj(Obj, PrpPthAy)
Next
End Function

Private Sub Z_DrszItrpp()
'BrwDrs DrszItrPP(Excel.Application.Addins, "Name Installed IsOpen FullName CLSId ")
'BrwDrs DrszItrPP(Fds(Db(SampFbzDutyDta), "Permit"), "Name Type Required")
'BrwDrs ItrPrpDrs(Application.VBE.VBProjects, "Name Type")
BrwDrs DrszItrpp(CurPj.VBComponents, "Name Type CmpTy=ShpCmpTy(Type)")
End Sub

Private Sub Z()
Z_DrszItrpp
MDta_Prp:
End Sub

