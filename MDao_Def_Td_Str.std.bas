Attribute VB_Name = "MDao_Def_Td_Str"
Option Explicit
Const CMod$ = "MDao_Td_Str."

Function TdStr$(A As Dao.TableDef)
Dim T$, Id$, Sk$, Rst$
'    Dim RstAy$(): RstAy = AyeEle(AyMinus(Fny, SkFny), Id)
'    Dim UseStar$(): UseStar = AyRpl(RstAy, "*", T)
'    Rst = StrApp(JnSpc(UseStar))
'TdStrTd = T & Id & Sk & Rst
'See also Stru$.  How should they
End Function

Function FnyzTdLy(TdLy$()) As String()
Dim O$(), TdStr
For Each TdStr In Itr(TdLy)
    PushIAy O, FnyzTdLin(TdStr)
Next
FnyzTdLy = AywDistSy(O)
End Function

Function FdStrAyzTdStr(TdStr) As String()
Dim F
For Each F In FnyzTdLin(TdStr)
    PushI FdStrAyzTdStr, FdStr(FdzStd(F))
Next
End Function
Function SampTdStr$()

End Function

Function TdStrz$(A As Database, T)
TdStrz = TdStr(A.TableDefs(T))
End Function

Function FnyzTdLin(TdLin) As String()
Dim T$, Rst$
AsgTRst TdLin, T, Rst
If HasSfx(T, "*") Then
    T = RmvSfx(T, "*")
    Rst = T & "Id " & Rst
End If
Rst = Replace(Rst, "*", T)
Rst = Replace(Rst, "|", " ")
FnyzTdLin = SySsl(Rst)
End Function

Function SkFnyzTdLin(A) As String()
Dim A1$, T$, Rst$
    A1 = TakBef(A, "|")
    If A1 = "" Then Exit Function
AsgTRst A1, T, Rst
T = RmvSfx(T, "*")
Rst = Replace(Rst, "*", T)
SkFnyzTdLin = SySsl(Rst)
End Function

Private Sub ZZ()
Dim A As Dao.TableDef
Dim B$()
'FnyzTdLin C
'SkFnyzTdLin C
End Sub

Private Sub Z()
End Sub

