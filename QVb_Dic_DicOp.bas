Attribute VB_Name = "QVb_Dic_DicOp"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Dic."

Function CvDic(A) As Dictionary
Set CvDic = A
End Function

Function AsetzDicKey(A As Dictionary) As Aset
Set AsetzDicKey = AsetzItr(A.Keys)
End Function

Function CvDicAy(A) As Dictionary()
CvDicAy = A
End Function

Function AddDicAy(A As Dictionary, Dy() As Dictionary) As Dictionary
Set AddDicAy = CloneDic(A)
Dim J%
For J = 0 To UB(Dy)
   PushDic AddDicAy, Dy(J)
Next
End Function

Function IupDic(A As Dictionary, By As Dictionary) As Dictionary 'Return New dictionary from A-Dic by Ins-or-upd By-Dic.  Ins: if By-Dic has key and A-Dic. _
Upd: K fnd in both, A-Dic-Val will be replaced by By-Dic-Val
Dim O As New Dictionary, K
For Each K In A.Keys
    If By.Exists(K) Then
        O.Add K, By(K)
    Else
        O(K) = By(K)
    End If
Next
Set IupDic = O
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
    If Not IsNm(CStr(K)) Then Exit Function
Next
DicAllKeyIsNm = True
End Function

Function DicAddKeyPfx(A As Dictionary, KeyPfx$) As Dictionary
Set DicAddKeyPfx = New Dictionary
Dim K
For Each K In A.Keys
    DicAddKeyPfx.Add KeyPfx & K, A(K)
Next
End Function
Function DicAyKy(A() As Dictionary) As Variant()
Dim I
For Each I In Itr(A)
   PushNoDupAy DicAyKy, CvDic(I).Keys
Next
End Function

Function DiczDry_TwoCol(Dry(), Optional Sep$ = " ") As Dictionary
Dim O As New Dictionary
If Si(Dry) <> 0 Then
   Dim Dr
   For Each Dr In Dry
       O.Add Dr(0), Dr(1)
   Next
End If
Set DiczDry_TwoCol = O
End Function

Function CloneDic(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, A(K)
Next
Set CloneDic = O
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
Function DryzDotAy(DotAy$()) As Variant()
Dim I, Lin
For Each I In Itr(DotAy)
    Lin = I
    PushI DryzDotAy, SplitDot(Lin)
Next
End Function
Function DryzDic(A As Dictionary, Optional InclDicValOptTy As Boolean) As Variant()
Dim I, Dr
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Si(K) = 0 Then Exit Function
For Each I In K
    If InclDicValOptTy Then
        Dr = Array(I, A(I), TypeName(A(I)))
    Else
        Dr = Array(I, A(I))
    End If
    Push DryzDic, Dr
Next
End Function

Function DicIntersectAy(A As Dictionary, B As Dictionary) As Dictionary
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
X: Set DicIntersectAy = O
End Function
Sub ThwIf_DifDic(A As Dictionary, B As Dictionary, Fun$, Optional N1$ = "A", Optional N2$ = "B")
If Not IsEqDic(A, B) Then Thw Fun, "2 given dic are diff", FmtQQ("[?] [?]", N1, N2), FmtDic(A), FmtDic(B)
End Sub

Function KeySet(A As Dictionary) As Aset
Set KeySet = AsetzItr(A.Keys)
End Function

Function KeySyzDic(A As Dictionary) As String()
KeySyzDic = SyzAy(A.Keys)
End Function

Function ValzDicIfKyJn$(A As Dictionary, Ky, Optional Sep$ = vbCrLf & vbCrLf)
Dim O$(), K
For Each K In Itr(Ky)
    If A.Exists(K) Then
        PushI O, A(K)
    End If
Next
ValzDicIfKyJn = Join(O, Sep)
End Function

Function SyzDicKy(Dic As Dictionary, Ky$()) As String()
Const CSub$ = CMod & "SyzDicKy"
Dim K
For Each K In Itr(Ky)
    If Dic.Exists(K) Then Thw CSub, "K of Ky not in Dic", "K Ky Dic", K, Ky, Dic
    PushI SyzDicKy, Dic(K)
Next
End Function

Function LineszDic$(A As Dictionary)
LineszDic = JnCrLf(FmtDic2(A))
End Function

Function FmtDic2(A As Dictionary) As String()
Dim K
For Each K In A.Keys
    Push FmtDic2, LyzKzLines(CStr(K), A(K))
Next
End Function

Function LyzKzLines(K$, Lines$) As String()
Dim J&
Dim Ly$()
    Ly = SplitCrLf(Lines)
For J = 0 To UB(Ly)
    Dim Lin
        Lin = Ly(J)
        If FstChr(Lin) = " " Then Lin = "~" & RmvFstChr(Lin)
    Push LyzKzLines, K & " " & Lin
Next
End Function

Function MaxSizAyDic%(AyDic As Dictionary) ' MthDic is DicOf_Mthn_zz_MthLinesAy
'MaxCntgMth is max-of-#-of-method per Mthn
Dim O%, K
For Each K In AyDic.Items
    O = Max(O, Si(AyDic(K)))
Next
MaxSizAyDic = O
End Function

Function MgeDic(A As Dictionary, PfxSsl$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
   Ny = SyzSS(PfxSsl)
   Ny = AddSfxzAy(Ny, "@")
If Si(Av) <> Si(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
   Dim J%
   For J = 0 To UB(Ny)
       Set D = Av(J)
       Push Dy, AddDicKeyPfx(A, Ny(J))
   Next
Set MgeDic = AddDicAy(A, Dy)
End Function

Function MinusDic(A As Dictionary, B As Dictionary) As Dictionary
If A.Count = 0 Then Set MinusDic = New Dictionary: Exit Function
If B.Count = 0 Then Set MinusDic = CloneDic(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set MinusDic = O
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

Function KeyzLikAyDic_Itm$(Dic As Dictionary, Itm$)
Dim K, LikAy$()
For Each K In Dic.Keys
    LikAy = Dic(K)
    If HitLikAy(Itm, LikAy) Then
        KeyzLikAyDic_Itm = K
        Exit Function
    End If
Next
End Function

Function KeyzLikssDic_Itm$(A As Dictionary, Itm$)
Dim Likss$, K
For Each K In A
    Likss = A(K)
    If HitLikss(Itm, Likss) Then KeyzLikssDic_Itm = K: Exit Function
Next
End Function

Private Sub Z_MaxSizAyDic()
Dim D As Dictionary, M%
'Set D = PjMthDic(CPj)
M = MaxSizAyDic(D)
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
AddDicAy B, C
AddDicKeyPfx B, A
DicAddOrUpd B, D, A, D
DicAllKeyIsNm B
DicAyKy C
IsEmpDic B
ThwIf_DifDic B, B, D, D, D
IsDicOfLines B
IsDicOfStr B
ValzDicIfKyJn B, A, D
SyzDicKy B, E
FmtDicTit B, D
LineszDic B
FmtDic2 B
MaxSizAyDic B
MgeDic B, D, G
MinusDic B, B
DicSelIntozAy B, E
DicSelIntoSy B, E
SyzDicKey B
DiczSwapKV B
DicTy B
End Sub

Function WbzNmzDiLines(NmzDiLines As Dictionary) As Workbook 'Assume each dic keys is name and each value is lines. _
create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Dim A As Dictionary: Set A = NmzDiLines
Ass IsItrOfNm(A.Keys)
Ass IsItrOfStr(A.Items)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook: Set O = NewWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = SqvzLines(A(K))
Next
X: Set WbzNmzDiLines = O
End Function


Function AzDiC(AzDiB As Dictionary, BzDiC As Dictionary) As Dictionary
Dim A, B, C
Set AzDiC = New Dictionary
For Each A In AzDiB.Keys
    B = AzDiB(A)
    If Not BzDiC.Exists(B) Then Thw CSub, "BzDiC does not contain B", "A B AzDiB BzDiC", A, B, AzDiB, BzDiC
    C = BzDiC(B)
    AzDiC.Add A, C
Next
End Function

Function DicAzDifVal(A As Dictionary, B As Dictionary) As Dictionary
Set DicAzDifVal = New Dictionary
Dim K, V
For Each K In A.Keys
    If B.Exists(K) Then
        V = A(K)
        If V <> B(K) Then DicAzDifVal.Add K, V
    End If
Next
End Function
Sub SetKv(O As Dictionary, K, V)
If O.Exists(K) Then
    Asg V, O(K)
Else
    O.Add K, V
End If
End Sub


Sub PushDic(O As Dictionary, A As Dictionary)
If IsNothing(O) Then
    Set O = CloneDic(A)
    Exit Sub
End If
Dim K
For Each K In A.Keys
    If O.Exists(A) Then Thw CSub, "O already has K.  Cannot push Dic-A to Dic-O", "K Dic-O Dic-A", K, O, A
    O.Add K, A(K)
Next
End Sub


Sub PushItmzSyDic(A As Dictionary, K, Itm)
Dim M$()
If A.Exists(K) Then
    M = A(K)
    PushI M, Itm
    A(K) = M
Else
    A.Add K, Sy(Itm)
End If
End Sub

Sub ThwNotSyDic(A As Dictionary, Fun$)
If Not IsDicOfSy(A) Then Thw Fun, "Given dictionary is not SyDic, all key is string and val is Sy", "Give-Dictionary", FmtDic(A)
End Sub


