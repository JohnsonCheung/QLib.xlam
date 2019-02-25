Attribute VB_Name = "MTp_Tp_Lin_Cln"
Option Explicit
Function ClnBrk1(A$(), Ny0) As Variant()
Dim O(), U%, Ny$(), L, T1$, T2$, NmDic As Dictionary, Ix%, Er$()
Ny = CvNy(Ny0)
U = UB(Ny)
ReDim O(U)
'O = AyMap(O, "EmpSy")
Set NmDic = IxDiczAy(Ny)
For Each L In A
    AsgTRst LTrim(L), T1, T2
    If NmDic.Exists(T1) Then
        Ix = NmDic(T1)
        Push O(Ix), T2 '<----
    End If
Next
Push O, ClnT1Chk(A, Ny)
ClnBrk1 = O
End Function

Function ClnT1Chk(A$(), T1Ay0) As String()
Dim T1Ay$(), L, O$()
T1Ay = CvNy(T1Ay0)
For Each L In A
    If Not HasEle(T1Ay, T1(L)) Then Push O, L
Next
If Sz(O) > 0 Then
    O = AyAddPfx(AyQuoteSq(O), Space(4))
    O = AyInsItm(O, FmtQQ("Following lines have invalid T1.  Valid T1 are [?]", JnSpc(T1Ay)))
End If
ClnT1Chk = O
End Function

Function ClnLin$(Lin)
If IsEmp(Lin) Then Exit Function
If IsDotLin(Lin) Then Exit Function
If IsSngTermLin(Lin) Then Exit Function
If IsDDLin(Lin) Then Exit Function
ClnLin = TakBefDD(Lin)
End Function

Function ClnLy(Ly$()) As String()
Dim L, S$
For Each L In Itr(Ly)
    PushI ClnLy, L
Next
End Function

Function ClnLnxAy(Ly$()) As Lnx()
Dim O()  As Lnx, L$, J%
For J = 0 To UB(Ly)
    L = ClnLin(Ly(J))
    If L <> "" Then
        Dim M  As Lnx
        Set M = New Lnx
        M.Ix = J
        M.Lin = Ly(J)
        Push O, M
    End If
Next
ClnLnxAy = O
End Function
