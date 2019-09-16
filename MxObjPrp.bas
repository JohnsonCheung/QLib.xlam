Attribute VB_Name = "MxObjPrp"
Option Compare Text
Option Explicit
Const CLib$ = "QDta."
Const CMod$ = CLib & "MxObjPrp."

Function AddFml(A As Drs, NewFld$, FunNm$, PmAy$()) As Drs
Dim Dy(): Dy = A.Dy
If Si(Dy) = 0 Then AddFml = A: Exit Function
Dim Dr, U&, Ixy1&(), Av()
Ixy1 = Ixy(A.Fny, PmAy)
U = UB(A.Fny)
For Each Dr In Dy
    If UB(Dr) <> U Then Thw CSub, "Dr-Si is diff", "Dr-Si U", UB(Dr), U
    Av = AwIxy(Dr, Ixy1)
    Push Dr, RunAv(FunNm, Av)
Next
AddFml = Drs(AddEleS(A.Fny, NewFld), Dy)
End Function

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

Function DrszItrPP(Itr, PP$) As Drs
Dim P$(): P = SyzSS(PP)
    Dim Obj, Dy(): For Each Obj In Itr
        PushI Dy, DrzObjPny(Obj, P)
    Next
DrszItrPP = Drs(P, Dy)
End Function

Function DrszItrPrpPthSy(Itr, PrpPthSy$()) As Drs
DrszItrPrpPthSy = Drs(PrpPthSy, DyoItrPrpPthSy(Itr, PrpPthSy))
End Function

Function DrzObjPny(Obj, Pny$(), Optional ThwEr As EmThw) As Variant()
Dim P
For Each P In Pny
    Push DrzObjPny, Prp(Obj, P, ThwEr)
Next
End Function

Function DyoItrPrpPthSy(Itr, PrpPthSy$()) As Variant()
Dim Obj As Object
For Each Obj In Itr
    Push DyoItrPrpPthSy, DrzPrpPthAy(Obj, PrpPthSy)
Next
End Function

Private Function P_QuietEmp(Obj, P)
On Error Resume Next
Asg CallByName(Obj, P, VbGet), P_QuietEmp
End Function

Function PrpzP(Obj, Prp)
Asg CallByName(Obj, Prp, VbGet), PrpzP
End Function

Function Prp(Obj, PrpPth, Optional ThwEr As EmThw)
Const CSub$ = CMod & "Prp"
'ThwIf_Nothing Obj, CSub
On Error GoTo X
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim PrpSeg$(): PrpSeg = Split(PrpPth, ".")
Dim O
    Set O = Obj
    Dim U%: U = UB(PrpSeg)
    Dim J%: For J = 0 To U - 1     ' U-1 is to skip the last Pth-Seg
        Set O = PrpzP(O, PrpSeg(J)) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next
Asg PrpzP(O, PrpSeg(U)), Prp ' Last Prp may be non-object, so must use 'Asg'
Exit Function
X:
Dim E$: E = Err.Description
If ThwEr = EiThwEr Then
    Thw CSub, "Err", "Er ObjTy PrpPth", E, TypeName(Obj), PrpPth
End If
End Function

Function PrpzP1(Obj, P, Optional ThwEr As EmThw)
Select Case True
Case ThwEr = EiNoThw: Asg P_QuietEmp(Obj, P), PrpzP1
Case Else: Stop
End Select
End Function

Private Sub WAsg3PP(PP_with_NewFldEqQteFmFld$, OPPzPrp$(), OPPzFml$(), OPPzAll$())
Dim I, S$
For Each I In SyzSS(PP_with_NewFldEqQteFmFld)
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

Private Sub Z()
MDta_Prp:
End Sub

Private Sub Z_DrszItrPrpPthSy()
'BrwDrs DrszItrPP(Excel.Application.Addins, "Name Installed IsOpen FullName CLSId ")
'BrwDrs DrszItrPP(Fds(Db(SampFbzDutyDta), "Permit"), "Name Type Required")
'BrwDrs ItrPrpDrs(Application.VBE.VBProjects, "Name Type")
BrwDrs DrszItrPrpPthSy(CPj.VBComponents, SyzSS("Name Type CmpTy=ShpCmpTy(Type)"))
End Sub