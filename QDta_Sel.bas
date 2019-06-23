Attribute VB_Name = "QDta_Sel"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDta_Sel."
Private Const Asm$ = "QDta"

Function DyoSel(Dy(), Ixy&()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DyoSel, AwIxy(Dr, Ixy)
Next
End Function
Function DyoSelAlwE(Dy(), Ixy&()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI DyoSelAlwE, AwIxyAlwE(Dr, Ixy)
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
    A = AwLik(Ay, Lik)
    If Si(A) = 0 Then
        PushI ExpandLikAy, Lik
    Else
        PushIAy ExpandLikAy, A
    End If
Next
End Function
Function LDrszJn(A As Drs, B As Drs, Jn$, Add$, Optional AnyFld$) As Drs
LDrszJn = JnDrs(A, B, Jn, Add, IsLeftJn:=True, AnyFld:=AnyFld)
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
Erase OFnyA, OFnyB
Dim F: For Each F In SyzSS(FFWiColon)
    With BrkBoth(F, ":")
        PushI OFnyA, .S1
        PushI OFnyB, .S2
    End With
Next
End Sub
Function GpRxy(Dy()) As Variant()
'Fm Dy : all col in @Dy will be used to gp
'Ret    : N-Gp of Rxy (Rec-Ix-Ay) pointing to @Dy.  That means each gp contain a Rxy.
'         and each ele in each Rxy is a Rix pointing a dist rec of @Dy
Dim K(), Dr, O(), Rix&: For Each Dr In Itr(Dy)
    Dim Gix&: Gix = IxzDyDr(K, Dr)
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

Function AddColzFst(D As Drs, Gpcc$) As Drs
'Fm D    : ..@Gpcc.. ! a drs with col-@Gpcc
'Fm Gpcc :           ! col-@Gpcc in @D have dup.
'Ret     : @D Fst    ! a drs of col-Fst add to @D at end.  col-Fst is bool value.  TRUE when if it fst rec of a gp
'                    ! and rst of rec of the gp to FALSE
Dim O As Drs: O = AddCol(D, "Fst", False) ' Add col-Fst with val all FALSE
If NoReczDrs(D) Then AddColzFst = O: Exit Function
Dim GDy(): GDy = DrszSel(D, Gpcc).Dy  ' Dy with Gp-col only.
Dim R(): R = GpRxy(GDy)                 ' Gp the @GDy into `GpRxy`
Dim Cix&: Cix = UB(O.Dy(0))             ' Las col Ix aft adding col-Fst
Dim Rxy: For Each Rxy In R               ' for each gp, get the Row-ixy (pointing to @D.Dy)
    Dim Rix&: Rix = Rxy(0)               ' Rix is Row-ix pointing one of @D.Dy which is the fst rec of a gp
    O.Dy(Rix)(Cix) = True
Next
AddColzFst = O
End Function

Function JnDrs(A As Drs, B As Drs, Jn$, Add$, Optional IsLeftJn As Boolean, Optional AnyFld$) As Drs
'Fm A        : ..@Jn-LHS..              ! It is a drs with col-@Jn-LHS.
'Fm B        : ..@Jn-RHS..@Add-RHS      ! It is a drs with col-@Jn-RHS & col-@Add-RHS.
'Fm Jn       : :SS-of-:ColonTerm        ! It is :SS-of-:ColTerm. :ColTerm: is a :Term with 1-or-0 [:]. :Term: is a fm :TLin: or :TermLin:  LHS of [:] is for @A and RHS of [:] is for @B
'                                       ! It is used to jn @A & @B
'Fm Add      : SS-of-ColonStr-Fld-@B    ! What col in @B to be added to @A.  It may use new name, if it has colon.
'Fm IsLeftJn :                          ! Is it left join, otherwise, it is inner join
'Fm AnyFld   : Fldn                     ! It is optional fld to be add to rslt drs stating if any rec in @B according to @Jn.
'                                       ! It is vdt only when IsLeftJn=True.
'                                       ! It has bool value.  It will be TRUE if @B has jn rec else FALSE.
'Ret         : ..@A..@Add-RHS..@AnyFld ! It has all fld from @A and @Add-RHS-fld and optional @AnyFld.
'                                       ! If @IsLeftJn, it will have at least same rec as @A, and may have if there is dup rec in @B accord to @Jn fld.
'                                       ! If not @IsLeftJn, only those records fnd in both @A & @B
Dim JnFnyA$(), JnFnyB$()
Dim AddFnyFm$(), AddFnyAs$()
    AsgFnyAB Jn, JnFnyA, JnFnyB
    AsgFnyAB Add, AddFnyFm, AddFnyAs
    
Dim AddIxy&(): AddIxy = IxyzSubAy(B.Fny, AddFnyFm, ThwNFnd:=True)
Dim BJnIxy&(): BJnIxy = IxyzSubAy(B.Fny, JnFnyB, ThwNFnd:=True)
Dim AJnIxy&(): AJnIxy = IxyzSubAy(A.Fny, JnFnyA, ThwNFnd:=True)

Dim Emp() ' it is for LeftJn and for those rec when @B has no rec joined.  It is for @Add-fld & @AnyFld.
          ' It has sam ele as @Add.  1 more fld is @AnyFld<>""
    If IsLeftJn Then
        ReDim Emp(UB(AddFnyFm))
        If AnyFld <> "" Then PushI Emp, False
    End If
Dim ODy()                       ' Bld %ODy for each %ADr, that mean fld-Add & fld-Any
    Dim ADr: For Each ADr In Itr(A.Dy)
        Dim JnVy():            JnVy = AwIxy(ADr, AJnIxy)                     'JnFld-Vy-Fm-@A
        Dim BDy():            BDy = DywKeySel(B.Dy, BJnIxy, JnVy, AddIxy) '@B-Dy-joined
        Dim NoRec As Boolean: NoRec = Si(BDy) = 0                           'no rec joined
            
        Select Case True
        Case NoRec And IsLeftJn: PushI ODy, AyzAdd(ADr, Emp) '<== ODy, Only for NoRec & LeftJn
        Case NoRec
        Case Else
            '
            Dim BDr: For Each BDr In BDy
                If AnyFld <> "" Then
                    Push BDr, True
                End If
                PushI ODy, AyzAdd(ADr, BDr) '<== ODy, for each %BDr in %BDy, push to %ODy
            Next
        End Select
    Next ADr

Dim O As Drs: O = Drs(SyNB(A.Fny, AddFnyAs, AnyFld), ODy)
JnDrs = O

If False Then
    Erase XX
    XBox "Debug JnDrs"
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
End Function
Function DywKeySel(Dy(), KeyIxy&(), Key(), SelIxy&()) As Variant()
DywKeySel = SelDy(DywKey(Dy, KeyIxy, Key), SelIxy)
End Function
Function SelDy(Dy(), SelIxy&()) As Variant()
'Ret : SubSet-of-col of @Dy indicated by @SelIxy
Dim Dr: For Each Dr In Itr(Dy)
    PushI SelDy, AwIxy(Dr, SelIxy)
Next
End Function

Function DywKey(Dy(), KeyIxy&(), Key()) As Variant()
'Ret : SubSet-of-row of @Dy for each row has val of %CurKey = @Key
Dim Dr: For Each Dr In Itr(Dy)
    Dim CurK: CurK = AwIxy(Dr, KeyIxy)
    If IsEqAy(CurK, Key) Then         '<- If %CurKey = @Key, select it.
        PushI DywKey, Dr
    End If
Next
End Function



Function InsColzDyVyBef(Dy(), Vy()) As Variant()
Dim Dr: For Each Dr In Itr(Dy)
    PushI InsColzDyVyBef, AyzAdd(Vy, Dr)
Next
End Function
Function InsColzDyBef(Dy(), V) As Variant()
InsColzDyBef = InsColzDyVyBef(Dy, Av(V))
End Function
Function InsColzDrsCC(A As Drs, CC$, V1, V2) As Drs
InsColzDrsCC = DwInsFF(A, CC, InsColzDyV2(A.Dy, V1, V2))
End Function
Function InsColzDrsC3(A As Drs, CCC$, V1, V2, V3) As Drs
InsColzDrsC3 = DwInsFF(A, CCC, InsColzDyV3(A.Dy, V1, V2, V3))
End Function
Function InsColzFront(A As Drs, C$, V) As Drs
InsColzFront = DwInsFF(A, C, InsColzDyBef(A.Dy, V))
End Function
Function InsCol(A As Drs, C$, V) As Drs
InsCol = InsColzFront(A, C, V)
End Function
Function UpdDrs(A As Drs, B As Drs, Jn$, Upd$, IsLefJn As Boolean) As Drs
'Fm  A  : ..@Jn-LHS..@Upd-LHS.. ! to be updated
'Fm  B  : ..@Jn-RHS..@Upd-RHS.. ! used to update @A.@Upd-LHS
'Fm  Jn : :SS-JnTerm            ! :JnTerm is :ColonTerm.  LHS is @A-fld and RHS is @B-fld
'Fm Upd : :Upd-UpdTerm          ! :UpdTer: is :ColTerm.  LHS is @A-fld and RHS is @B-fld
'Ret    : sam as @A             ! new Drs from @A with @A.@Upd-LHS updated from @B.@Upd-RHS. @@
Dim C As Dictionary: Set C = DiczDrsCC(B)
Dim O As Drs
    O.Fny = A.Fny
    Dim Dr, K
    For Each Dr In A.Dy
        K = Dr(0)
        If C.Exists(K) Then
            Dr(0) = C(K)
        End If
        PushI O.Dy, Dr
    Next
UpdDrs = O
'BrwDrs3 A, B, O, NN:="A B O", Tit:= _
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
    Dim Dr: For Each Dr In Itr(A.Dy)
        Dim Ix&: Ix = IxzDyDr(OKey, Dr)
        If Ix = -1 Then
            PushI OCnt, 1
            PushI OKey, Dr
        Else
            OCnt(Ix) = OCnt(Ix) + 1
        End If
    Next
Dim ODy()
    Dim J&: For Each Dr In Itr(OKey)
        Push Dr, OCnt(J)
        PushI ODy, Dr
        J = J + 1
    Next
SelDist = DrszFF(FF & " Cnt", ODy)
End Function

Function DrszSel(A As Drs, FF$) As Drs
DrszSel = DrszSelFny(A, SyzSS(FF))
End Function

Function DrszSelFny(A As Drs, Fny$()) As Drs
ThwNotSuperAy A.Fny, Fny
DrszSelFny = Drs(Fny, DyoSel(A.Dy, Ixy(A.Fny, Fny)))
End Function

Function DrszSelAs(A As Drs, FFAs$) As Drs
Dim FA$(), Fb$(): AsgFnyAB FFAs, FA, Fb
DrszSelAs = Drs(Fb, DrszSelFny(A, FA).Dy)
End Function

Function DrszSelAlwEzFny(A As Drs, Fny$()) As Drs
If IsEqAy(A.Fny, Fny) Then DrszSelAlwEzFny = A: Exit Function
DrszSelAlwEzFny = Drs(Fny, DyoSelAlwE(A.Dy, IxyzAlwE(A.Fny, Fny)))
End Function

Function DrszSelAlwE(A As Drs, FF$) As Drs
DrszSelAlwE = DrszSelAlwEzFny(A, SyzSS(FF))
End Function

Function DrszUpdC(A As Drs, C$, V) As Drs
Dim I&: I = IxzAy(A.Fny, C)
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    Dr(I) = V
    PushI Dy, Dr
Next
DrszUpdC = Drs(A.Fny, Dy)
End Function

Function DrszUpdCC(A As Drs, CC$, V1, V2) As Drs
Dim I1&, I2&: AsgIx A, CC, I1, I2
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    Dr(I1) = V1
    Dr(I2) = V2
    PushI Dy, Dr
Next
DrszUpdCC = Drs(A.Fny, Dy)
End Function

Function SelDt(A As Dt, FF$) As Dt
SelDt = DtzDrs(DrszSel(DrszDt(A), FF), A.DtNm)
End Function


Private Sub Z()
MDta_Sel:
End Sub
