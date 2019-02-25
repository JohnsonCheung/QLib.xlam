Attribute VB_Name = "MVb_Dic"
Option Explicit
Const CMod$ = "MVb_Dic."

Function CvDic(A) As Dictionary
Set CvDic = A
End Function

Function AsetzDicKey(A As Dictionary) As Aset
Set AsetzDicKey = AsetzItr(A.Keys)
End Function

Function CvDicAy(A) As Dictionary()
CvDicAy = A
End Function

Function DicAyAdd(A As Dictionary, Dy() As Dictionary) As Dictionary
Set DicAyAdd = DicClone(A)
Dim J%
For J = 0 To UB(Dy)
   PushDic DicAyAdd, Dy(J)
Next
End Function

Function AddDicKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
Set AddDicKeyPfx = O
End Function

Sub DicAddOrUpd(A As Dictionary, K$, V, Sep$)
If A.Exists(K) Then
    A(K) = A(K) & Sep & V
Else
    A.Add K, V
End If
End Sub

Function DicAllKeyIsNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
DicAllKeyIsNm = True
End Function

Function DicAyKy(A() As Dictionary) As Variant()
Dim I
For Each I In Itr(A)
   PushNoDupAy DicAyKy, CvDic(I).Keys
Next
End Function

Function DiczDryOfTwoCol(Dry(), Optional Sep$ = " ") As Dictionary
Dim O As New Dictionary
If Sz(Dry) <> 0 Then
   Dim Dr
   For Each Dr In Dry
       O.Add Dr(0), Dr(1)
   Next
End If
Set DiczDryOfTwoCol = O
End Function

Function DicClone(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, A(K)
Next
Set DicClone = O
End Function

Function DrDicKy(A As Dictionary, Ky$()) As Variant()
Dim O(), I, J&
ReDim O(UB(Ky))
For Each I In Ky
    If A.Exists(I) Then
        O(J) = A(I)
    End If
    J = J + 1
Next
DrDicKy = O
End Function

Function DicFny(InclDicValOptTy As Boolean) As String()
DicFny = SplitSpc("Key Val"): If InclDicValOptTy Then PushI DicFny, "ValTy"
End Function

Function DryzDic(A As Dictionary, Optional InclDicValOptTy As Boolean) As Variant()
Dim I, Dr
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Sz(K) = 0 Then Exit Function
For Each I In K
    If InclDicValOptTy Then
        Dr = Array(I, A(I), TypeName(A(I)))
    Else
        Dr = Array(I, A(I))
    End If
    Push DryzDic, Dr
Next
End Function

Function DicIntersect(A As Dictionary, B As Dictionary) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
If B.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set DicIntersect = O
End Function


Sub ThwDifDic(A As Dictionary, B As Dictionary, Fun$, Optional N1$ = "A", Optional N2$ = "B")
If Not IsEqDic(A, B) Then Thw Fun, "2 given dic are diff", FmtQQ("[?] [?]", N1, N2), FmtDic(A), FmtDic(B)
End Sub


Function DicKeySy(A As Dictionary) As String()
DicKeySy = SyzAy(A.Keys)
End Function

Function DicKyJnVal$(A As Dictionary, Ky, Optional Sep$ = vbCrLf & vbCrLf)
Dim O$(), K
For Each K In Itr(Ky)
    If A.Exists(K) Then
        PushI O, A(K)
    End If
Next
DicKyJnVal = Join(O, Sep)
End Function

Function SyDicKy(Dic As Dictionary, Ky$()) As String()
Const CSub$ = CMod & "SyDicKy"
Dim K
For Each K In Itr(Ky)
    If Dic.Exists(K) Then Thw CSub, "K of Ky not in Dic", "K Ky Dic", K, Ky, Dic
    PushI SyDicKy, Dic(K)
Next
End Function

Function DicLblLy(A As Dictionary, Lbl$) As String()
PushI DicLblLy, Lbl
PushI DicLblLy, vbTab & "Count=" & A.Count
PushIAy DicLblLy, AyAddPfx(FmtDic(A, InclValTy:=True), vbTab)
End Function

Function LinesDic(A As Dictionary) As String
LinesDic = JnCrLf(FmtDic2(A))
End Function

Function FmtDic2(A As Dictionary) As String()
Dim O$(), K
If A.Count = 0 Then Exit Function
For Each K In A.Keys
    Push O, FmtDic2__1(CStr(K), A(K))
Next
FmtDic2 = O
End Function

Function FmtDic2__1(K$, Lines$) As String()
Dim O$(), J&
Dim Ly$()
    Ly = SplitCrLf(Lines)
For J = 0 To UB(Ly)
    Dim Lin$
        Lin = Ly(J)
        If FstChr(Lin) = " " Then Lin = "~" & RmvFstChr(Lin)
    Push O, K & " " & Lin
Next
FmtDic2__1 = O
End Function

Function DicMap(A As Dictionary, ValMapFun$) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, Run(ValMapFun, A(K))
Next
Set DicMap = O
End Function

Function DicMaxValSz%(A As Dictionary)
'MthDic is DicOf_MthNm_zz_MthLinesAy
'MaxMthCnt is max-of-#-of-method per MthNm
Dim O%, K
For Each K In A.Keys
    O = Max(O, Sz(A(K)))
Next
DicMaxValSz = O
End Function

Function DicMge(A As Dictionary, PfxSsl$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
   Ny = SySsl(PfxSsl)
   Ny = AyAddSfx(Ny, "@")
If Sz(Av) <> Sz(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
   Dim J%
   For J = 0 To UB(Ny)
       Set D = Av(J)
       Push Dy, AddDicKeyPfx(A, Ny(J))
   Next
Set DicMge = DicAyAdd(A, Dy)
End Function

Function DicMinus(A As Dictionary, B As Dictionary) As Dictionary
If A.Count = 0 Then Set DicMinus = New Dictionary: Exit Function
If B.Count = 0 Then Set DicMinus = DicClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set DicMinus = O
End Function

Function DicSelIntozAy(A As Dictionary, Ky$()) As Variant()
Dim O()
Dim U&: U = UB(Ky)
ReDim O(U)
Dim J&
For J = 0 To U
   If Not A.Exists(Ky(J)) Then Stop
   O(J) = A(Ky(J))
Next
DicSelIntozAy = O
End Function

Function DicSelIntoSy(A As Dictionary, Ky$()) As String()
DicSelIntoSy = SyzAy(DicSelIntozAy(A, Ky))
End Function

Function SyzDicKey(A As Dictionary) As String()
SyzDicKey = SyzAy(A.Keys)
End Function

Function DiczSwapKV(A As Dictionary) As Dictionary
Dim K
Set DiczSwapKV = New Dictionary
For Each K In A.Keys
    DiczSwapKV.Add A(K), K
Next
End Function

Function DicValOpt(A As Dictionary, K)
If IsNothing(A) Then Exit Function
If A.Exists(K) Then Asg A(K), DicValOpt
End Function

Function Keyz_LikAyDic_Itm$(Dic As Dictionary, Itm)
Dim K, LikAy$()
For Each K In Dic.Keys
    LikAy = Dic(K)
    If HitLikAy(Itm, LikAy) Then
        Keyz_LikAyDic_Itm = K
        Exit Function
    End If
Next
End Function


Function Keyz_LikssDic_Itm$(A As Dictionary, Itm)
Dim Likss$, K
For Each K In A
    Likss = A(K)
    If HitLikss(Itm, Likss) Then Keyz_LikssDic_Itm = K: Exit Function
Next
End Function

Private Sub Z_DicMaxValSz()
Dim D As Dictionary, M%
'Set D = PjMthDic(CurPj)
M = DicMaxValSz(D)
Stop
End Sub

Private Sub ZZ()
Dim A As Variant
Dim B As Dictionary
Dim C() As Dictionary
Dim D$
Dim E$()
Dim F As Boolean
Dim G()
CvDic A
CvDicAy A
DicAyAdd B, C
AddDicKeyPfx B, A
DicAddOrUpd B, D, A, D
DicAllKeyIsNm B
DicAyKy C
IsDiczEmp B
ThwDifDic B, B, D, D, D
IsDiczLines B
IsDiczStr B
DicKyJnVal B, A, D
SyDicKy B, E
DicLblLy B, D
LinesDic B
FmtDic2 B
FmtDic2__1 D, D
DicMap B, D
DicMaxValSz B
DicMge B, D, G
DicMinus B, B
DicSelIntozAy B, E
DicSelIntoSy B, E
SyzDicKey B
DiczSwapKV B
DicTy B
Keyz_LikssDic_Itm B, A
End Sub

Private Sub Z()
Z_DicMaxValSz
End Sub
Function WbzNmToLinesDic(A As Dictionary) As Workbook
'Assume each dic keys is name and each value is lines
'Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Ass IsItrzNm(A.Keys)
Ass IsItrzStr(A.Items)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook
Set O = NewWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = VSqLines(A(K))
Next
X: Set WbzNmToLinesDic = O
End Function


