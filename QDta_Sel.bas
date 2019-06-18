Attribute VB_Name = "QDta_Sel"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Sel."
Private Const Asm$ = "QDta"

Function DryzSel(Dry(), Ixy&()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI DryzSel, AywIxy(Drv, Ixy)
Next
End Function
Function DryzSelAlwE(Dry(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DryzSelAlwE, AywIxyAlwE(Dr, Ixy)
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
Function LDrszJn(A As Drs, B As Drs, Jn$, Add$, Optional AnyFld$) As Drs
LDrszJn = DrszJn(A, B, Jn, Add, IsLeftJn:=True, AnyFld:=AnyFld)
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
Function GpRxy(Dry()) As Variant()
'Fm Dry : all col in @Dry will be used to gp
'Ret    : N-Gp of Rxy (Rec-Ix-Ay) pointing to @Dry.  That means each gp contain a Rxy.
'         and each ele in each Rxy is a Rix pointing a dist rec of @Dry
Dim K(), Dr, O(), Rix&: For Each Dr In Itr(Dry)
    Dim Gix&: Gix = IxzDryDr(K, Dr)
    If Gix = -1 Then
        Dim Rxy&(): ReDim Rxy(0)
        Rxy(0) = Rix
        PushI O, Rxy      '<== Put Rix to Oup-O
        PushI K, Dr       '<-- Put Dr to K
    Else
        PushI O(Gix), Rix '<== Put Rix to Oup-O
    End If
    Rix = Rix + 1
Next
GpRxy = O
End Function

Function DrszAddFst(D As Drs, Gpcc$) As Drs
'Fm D    : ..@Gpcc.. ! a drs with col-@Gpcc
'Fm Gpcc :           ! col-@Gpcc in @D have dup.
'Ret     : @D Fst    ! a drs of col-Fst add to @D at end.  col-Fst is bool value.  TRUE when if it fst rec of a gp
'                    ! and rst of rec of the gp to FALSE
Dim O As Drs: O = DrszAddCV(D, "Fst", False) ' Add col-Fst with val all FALSE
Dim GDry(): GDry = DrszSel(D, Gpcc).Dry  ' Dry with Gp-col only.
Dim R(): R = GpRxy(GDry)                 ' Gp the @GDry into `GpRxy`
Dim Cix&: Cix = UB(O.Dry(0))             ' Las col Ix aft adding col-Fst
Dim Rxy: For Each Rxy In R               ' for each gp, get the Row-ixy (pointing to @D.Dry)
    Dim Rix&: Rix = Rxy(0)               ' Rix is Row-ix pointing one of @D.Dry which is the fst rec of a gp
    O.Dry(Rix)(Cix) = True
Next
DrszAddFst = O
End Function

Function DrszJn(A As Drs, B As Drs, Jn$, Add$, Optional IsLeftJn As Boolean, Optional AnyFld$) As Drs
'Fm A        : ..@Jn-LHS..              ! It is a drs with col-@Jn-LHS.
'Fm B        : ..@Jn-RHS..@Add-RHS      ! It is a drs with col-@Jn-RHS and col-@Add-RHS.
'Fm Jn       : SS-of-ColonStr           ! It is SS-str of str term with optional [:].  LHS of [:] is for @A and RHS of [:] is for @B
'                                       ! It is used to jn @A & @B
'Fm Add      : SS-of-ColonStr-Fld-@B    ! What col in @B to be added to @A.  It may use new name, if it has colon.
'Fm IsLeftJn :                          ! Is it left join, otherwise, it is inner join
'Fm AnyFld   : Fldn                     ! It is optional fld to be add to rslt drs stating if any rec in @B according to @Jn.
'                                       ! It is vdt only when IsLeftJn=True.
'                                       ! It has bool value.  It will be TRUE if @B has jn rec else FALSE.
'Ret         : ..@A..@Add-RHS.. @AnyFld ! It has all fld from @A and @Add-RHS-fld and optional @AnyFld.
'                                       ! If @IsLeftJn, it will have at least same rec as @A, and may have if there is dup rec in @B accord to @Jn fld.
'                                       ! If not @IsLeftJn, only those records fnd in both @A & @B
'
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
    Case NoRec And IsLeftJn And AnyFld = "": PushI ODry, AyzAdd(Dr, Emp)
    Case NoRec And IsLeftJn:                 PushI ODry, AyzAdd(Dr, EmpWithAny)
    Case NoRec
    Case AnyFld = ""
        For Each IDr In IDry
            PushI ODry, AyzAdd(Dr, IDr)
        Next
    Case Else
        For Each IDr In IDry
            PushI IDr, True
            PushI ODry, AyzAdd(Dr, IDr)
        Next
    End Select
Next
O = Drs(SyNonBlank(A.Fny, AddFnyAs, AnyFld), ODry)

If False Then
    Erase XX
    XBox "Debug DrszJn"
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
DrszJn = O
End Function

Function DrywIxyVySel(Dry(), WhIxy&(), Vy(), SelIxy&()) As Variant()
Dim Dr, IVy, IDr()
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
    PushI InsColzDryVyBef, AyzAdd(Vy, Dr)
Next
End Function
Function InsColzDryBef(Dry(), V) As Variant()
InsColzDryBef = InsColzDryVyBef(Dry, Av(V))
End Function
Function InsColzDrsCC(A As Drs, CC$, V1, V2) As Drs
InsColzDrsCC = DrszInsFF(A, CC, InsColzDryV2(A.Dry, V1, V2))
End Function
Function InsColzDrsC3(A As Drs, CCC$, V1, V2, V3) As Drs
InsColzDrsC3 = DrszInsFF(A, CCC, InsColzDryV3(A.Dry, V1, V2, V3))
End Function
Function InsColzFront(A As Drs, C$, V) As Drs
InsColzFront = DrszInsFF(A, C, InsColzDryBef(A.Dry, V))
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

Private Sub Z_SelDist()
BrwDrs SelDist(DMthP, "Mdn Ty")
End Sub

Function SelDist(D As Drs, FF$) As Drs
'Fm  D : ..{Gpcc} {C}.. ! it has columns-Gpcc and column-C
'Ret   : {Gpcc} {C}     ! where C is group of column-C @@
Dim OKey(), OCnt&()
    Dim A As Drs: A = DrszSel(D, FF)
    Dim I%: I = Si(A.Fny)
    Dim Dr: For Each Dr In Itr(A.Dry)
        Dim Ix&: Ix = IxzDryDr(OKey, Dr)
        If Ix = -1 Then
            PushI OCnt, 1
            PushI OKey, Dr
        Else
            OCnt(Ix) = OCnt(Ix) + 1
        End If
    Next
Dim ODry()
    Dim J&: For Each Dr In Itr(OKey)
        Push Dr, OCnt(J)
        PushI ODry, Dr
        J = J + 1
    Next
SelDist = DrszFF(FF & " Cnt", ODry)
End Function

Function DrszSel(A As Drs, FF$) As Drs
Dim Fny$(): Fny = SyzSS(FF)
DrszSel = Drs(Fny, DryzSel(A.Dry, Ixy(A.Fny, Fny)))
End Function

Function DrszSelzFny(A As Drs, Fny$()) As Drs
ThwNotSuperAy A.Fny, Fny
DrszSelzFny = Drs(Fny, DryzSel(A.Dry, Ixy(A.Fny, Fny)))
End Function

Function DrszSelAs(A As Drs, FFAs$) As Drs
Dim FA$(), Fb$(): AsgFnyAB FFAs, FA, Fb
DrszSelAs = Drs(Fb, DrszSelzFny(A, FA).Dry)
End Function

Function DrszSelAlwEzFny(A As Drs, Fny$()) As Drs
If IsEqAy(A.Fny, Fny) Then DrszSelAlwEzFny = A: Exit Function
DrszSelAlwEzFny = Drs(Fny, DryzSelAlwE(A.Dry, IxyzAlwE(A.Fny, Fny)))
End Function

Function DrszSelAlwE(A As Drs, FF$) As Drs
DrszSelAlwE = DrszSelAlwEzFny(A, SyzSS(FF))
End Function

Function DrszUpdC(A As Drs, C$, V) As Drs
Dim I&: I = IxzAy(A.Fny, C)
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    Dr(I) = V
    PushI Dry, Dr
Next
DrszUpdC = Drs(A.Fny, Dry)
End Function

Function DrszUpdCC(A As Drs, CC$, V1, V2) As Drs
Dim I1&, I2&: AsgIx A, CC, I1, I2
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    Dr(I1) = V1
    Dr(I2) = V2
    PushI Dry, Dr
Next
DrszUpdCC = Drs(A.Fny, Dry)
End Function

Function SelDt(A As Dt, FF$) As Dt
SelDt = DtzDrs(DrszSel(DrszDt(A), FF), A.DtNm)
End Function


Private Sub ZZ()
MDta_Sel:
End Sub
