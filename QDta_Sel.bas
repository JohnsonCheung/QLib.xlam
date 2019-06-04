Attribute VB_Name = "QDta_Sel"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Sel."
Private Const Asm$ = "QDta"

Function SelDry(Dry(), Ixy&()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI SelDry, AywIxy(Drv, Ixy)
Next
End Function
Function SelDryAlwEmp(Dry(), Ixy&()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI SelDryAlwEmp, AywIxy(Drv, Ixy)
Next
End Function
Function ExpandFF(FF$, Fny$()) As String() '
ExpandFF = ExpandLikAy(TermAy(FF), Fny)
End Function
Function ExpandLikAy(LikAy$(), Ay$()) As String() 'Put each expanded-ele in likAy to return a return ay. _
Expanded-ele means either the ele itself if there is no ele in Ay is like the `ele` _
                   or     the lik elements in Ay with the given `ele`
Dim Lik
For Each Lik In LikAy
    Dim A$()
    A = AywLik(Ay, Lik)
    If Si(A) = 0 Then
        PushI ExpandLikAy, Lik
    Else
        PushIAy ExpandLikAy, A
    End If
Next
End Function
Function LJnDrs(A As Drs, B As Drs, Jn$, Add$, Optional AnyFld$) As Drs
LJnDrs = JnDrs(A, B, Jn, Add, IsLeftJn:=True, AnyFld:=AnyFld)
End Function
Function JnDrs(A As Drs, B As Drs, Jn$, Add$, Optional IsLeftJn As Boolean, Optional AnyFld$) As Drs
Dim Dr, IDr, Dr1(), IDry(), ODry(), AddFny$(), AddFnyFm$(), AddFnyAs$(), F, JnFny$(), JnFnyA$(), JnFnyB$(), AJnIxy&(), BJnIxy&(), AddIxy&(), Vy()
Dim Emp(), EmpWithAny(), NoRec As Boolean, O As Drs
JnFny = SyzSS(Jn)
For Each F In JnFny
    With BrkBoth(F, ":")
        PushI JnFnyA, .S1
        PushI JnFnyB, .S2
    End With
Next
AddFny = SyzSS(Add)
For Each F In AddFny
    With BrkBoth(F, ":")
        PushI AddFnyFm, .S1
        PushI AddFnyAs, .S2
    End With
Next
AddIxy = IxyzSubAy(B.Fny, AddFnyFm)
BJnIxy = IxyzSubAy(B.Fny, JnFnyB)
AJnIxy = IxyzSubAy(A.Fny, JnFnyA)
If IsLeftJn Then ReDim Emp(UB(AddFny))
If IsLeftJn And AnyFld <> "" Then ReDim EmpWithAny(UB(AddFny)): PushI EmpWithAny, False
For Each Dr In Itr(A.Dry)
    Vy = AywIxy(Dr, AJnIxy)
    IDry = DrywIxyVySel(B.Dry, BJnIxy, Vy, AddIxy)
    NoRec = Si(IDry) = 0
    Select Case True
    Case NoRec And IsLeftJn And AnyFld = "": PushI ODry, AddAy(Dr, Emp)
    Case NoRec And IsLeftJn:                 PushI ODry, AddAy(Dr, EmpWithAny)
    Case NoRec
    Case AnyFld = ""
        For Each IDr In IDry
            PushI ODry, AddAy(Dr, IDr)
        Next
    Case Else
        For Each IDr In IDry
            PushI IDr, True
            PushI ODry, AddAy(Dr, IDr)
        Next
    End Select
Next
O = Drs(SyNonBlank(A.Fny, AddFnyAs, AnyFld), ODry)

If False Then
    Erase XX
    X "*****************"
    X "** Debug JnDrs **"
    X "*****************"
    X "A-Fny  : " & TermLin(A.Fny)
    X "B-Fny  : " & TermLin(B.Fny)
    X "Jn     : " & Jn
    X "Add    : " & Add
    X "IsLefJn: " & IsLeftJn
    X "AnyFld : [" & AnyFld & "]"
    X "O-Fny  : " & TermLin(O.Fny)
    X "More ..: A-Drs B-Drs Rslt"
    X LyzNmDrs("A-Drs  : ", A)
    X LyzNmDrs("B-Drs  : ", B)
    X LyzNmDrs("Rslt   : ", O)
    Brw XX
    Erase XX
    Stop
End If
JnDrs = O
End Function

Function DrywIxyVySel(Dry(), WhIxy&(), Vy(), SelIxy&()) As Variant()
Dim Dr, IVy(), IDr()
For Each Dr In Itr(Dry)
    IVy = AywIxy(Dr, WhIxy)
    If IsEqAy(Vy, IVy) Then
        IDr = AywIxy(Dr, SelIxy)
        PushI DrywIxyVySel, IDr
    End If
Next
End Function
Function InsColzDryVyBef(Dry(), Vy()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI InsColzDryVyBef, AddAy(Vy, Dr)
Next
End Function
Function InsColzDryBef(Dry(), V) As Variant()
InsColzDryBef = InsColzDryVyBef(Dry, Av(V))
End Function
Function InsColzDrsCCBef(A As Drs, CC$, V1, V2) As Drs
InsColzDrsCCBef = Drs(AddSy(SyzSS(CC), A.Fny), InsColzDryVyBef(A.Dry, Av(V1, V2)))
End Function
Function InsColzDrsBef(A As Drs, C$, V) As Drs
InsColzDrsBef = Drs(AddSy(Sy(C), A.Fny), InsColzDryBef(A.Dry, V))
End Function
Function UpdDrs(A As Drs, B As Drs) As Drs
'Fm  A      K X    ! to be updated
'Fm  B      K NewX ! used to update A.  K is unique
'Ret UpdDrs K X    ! new Drs from A with A.X may updated from B.NewX.

Dim C As Dictionary: Set C = DiczDrsCC(B)
Dim O As Drs
    O.Fny = A.Fny
    Dim Dr, K
    For Each Dr In A.Dry
        K = Dr(0)
        If C.Exists(K) Then
            Dr(0) = C(K)
        End If
        PushI O.Dry, Dr
    Next
UpdDrs = O
'BrwDrs3 A, B, O, NN:="A B O", Tit:= _
"Fm  A      K X    ! to be updated" & vbcrlf & _
"Fm  B      K NewX ! used to update A.  K is unique"  & vbcrlf & _
"Ret UpdDrs K X    ! new Drs from A with A.X may updated from B.NewX.
Stop
End Function
Function SelDrs(A As Drs, FF$) As Drs
Dim Fny$(): Fny = ExpandFF(FF, A.Fny)
ThwNotSuperAy A.Fny, Fny
SelDrs = SelDrsAlwEmpzFny(A, Fny)
End Function

Function SelDrszFny(A As Drs, Fny$()) As Drs
SelDrszFny = Drs(Fny, SelDry(A.Dry, Ixy(A.Fny, Fny)))
End Function

Function SelDrsAlwEmpzFny(A As Drs, Fny$()) As Drs
If IsEqAy(A.Fny, Fny) Then SelDrsAlwEmpzFny = A: Exit Function
SelDrsAlwEmpzFny = Drs(Fny, SelDry(A.Dry, Ixy(A.Fny, Fny)))
End Function

Function SelDrsAlwEmp(A As Drs, FF$) As Drs
SelDrsAlwEmp = SelDrsAlwEmpzFny(A, TermAy(FF))
End Function

Private Sub Z_SelDrs()
'BrwDrs SelDrs(Vmd.MthDrs, "Mthn Mdy Ty Mdn")
'BrwDrs Vmd.MthDrs
End Sub

Function SelDt(A As Dt, FF$) As Dt
SelDt = DtzDrs(SelDrs(DrszDt(A), FF), A.DtNm)
End Function


Private Sub ZZ()
Z_SelDrs
MDta_Sel:
End Sub
