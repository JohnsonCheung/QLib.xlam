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
Function FnyAzJn(Jn$) As String()
Dim J
For Each J In SyzSS(Jn)
    PushI FnyAzJn, BefOrAll(J, ":")
Next
End Function
Function FnyBzJn(Jn$) As String()
Dim J
For Each J In SyzSS(Jn)
    PushI FnyBzJn, AftOrAll(J, ":")
Next
End Function

Sub AsgFnyAB(FFWiColon$, OFnyA$(), OFnyB$())
Dim F
Erase OFnyA, OFnyB
For Each F In SyzSS(FFWiColon)
    With BrkBoth(F, ":")
        PushI OFnyA, .S1
        PushI OFnyB, .S2
    End With
Next
End Sub
Function JnDrs(A As Drs, B As Drs, Jn$, Add$, Optional IsLeftJn As Boolean, Optional AnyFld$) As Drs
Dim Dr, IDr, Dr1(), IDry(), ODry(), AddFny$(), AddFnyFm$(), AddFnyAs$(), F, JnFnyA$(), JnFnyB$(), AJnIxy&(), BJnIxy&(), AddIxy&(), Vy()
Dim Emp(), EmpWithAny(), NoRec As Boolean, O As Drs
AsgFnyAB Jn, JnFnyA, JnFnyB
AsgFnyAB Add, AddFnyFm, AddFnyAs
AddIxy = IxyzSubAy(B.Fny, AddFnyFm, ThwNFnd:=True)
BJnIxy = IxyzSubAy(B.Fny, JnFnyB, ThwNFnd:=True)
AJnIxy = IxyzSubAy(A.Fny, JnFnyA, ThwNFnd:=True)
If IsLeftJn Then ReDim Emp(UB(AddFnyFm))
If IsLeftJn And AnyFld <> "" Then ReDim EmpWithAny(UB(AddFnyFm)): PushI EmpWithAny, False
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
Function InsColzFront(A As Drs, C$, V) As Drs
InsColzFront = Drs(AddSy(Sy(C), A.Fny), InsColzDryBef(A.Dry, V))
End Function
Function InsCol(A As Drs, C$, V) As Drs
InsCol = InsColzFront(A, C, V)
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
'Dim Fny$(): Fny = ExpandFF(FF, A.Fny)
SelDrs = SelDrsAlwEmpzFny(A, SyzSS(FF))
End Function

Function SelDrszFny(A As Drs, Fny$()) As Drs
ThwNotSuperAy A.Fny, Fny
SelDrszFny = Drs(Fny, SelDry(A.Dry, Ixy(A.Fny, Fny)))
End Function

Function SelDrszAs(A As Drs, FFAs$) As Drs
Dim FA$(), FB$(): AsgFnyAB FFAs, FA, FB
SelDrszAs = Drs(FB, SelDrszFny(A, FA).Dry)
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
