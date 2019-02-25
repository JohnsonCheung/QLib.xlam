Attribute VB_Name = "MDao_Def_Td_Str"
Option Explicit
Const CMod$ = "MDao_Td_Str."

Function TdStr$(A As DAO.TableDef)
Dim T$, Id$, Sk$, Rst$
'    Dim RstAy$(): RstAy = AyeEle(AyMinus(Fny, SkFny), Id)
'    Dim UseStar$(): UseStar = AyRpl(RstAy, "*", T)
'    Rst = StrApp(JnSpc(UseStar))
'TdStrTd = T & Id & Sk & Rst
'See also Stru$.  How should they
End Function

Function FnyzTdStrAy(TdStrAy$()) As String()
Dim O$(), TdStr
For Each TdStr In Itr(TdStrAy)
    PushIAy O, FnyzTdStr(TdStr)
Next
FnyzTdStrAy = AywDistSy(O)
End Function

Function FdStrAyzTdStr(TdStr) As String()
Dim F
For Each F In FnyzTdStr(TdStr)
    PushI FdStrAyzTdStr, FdStr(StdFd(F))
Next
End Function
Function SampTdStr$()

End Function

Function TdStrz$(A As Database, T)
TdStrz = TdStr(A.TableDefs(T))
End Function
Function FnyzTdStr(TdStr) As String()
Dim T$, Rst$
AsgTRst TdStr, T, Rst
If HasSfx(T, "*") Then
    T = RmvSfx(T, "*")
    Rst = T & "Id " & Rst
End If
Rst = Replace(Rst, "*", T)
Rst = Replace(Rst, "|", " ")
FnyzTdStr = SySsl(Rst)
End Function

Function TdStrLines$(A As DAO.TableDef)
'TdStrLines = JnCrLf(TdStrLy(A))
End Function

Function TdStrSk(A) As String()
Dim A1$, T$, Rst$
    A1 = TakBef(A, "|")
    If A1 = "" Then Exit Function
AsgTRst A1, T, Rst
T = RmvSfx(T, "*")
Rst = Replace(Rst, "*", T)
TdStrSk = SySsl(Rst)
End Function

Private Sub ZZ()
Dim A As DAO.TableDef
Dim B$()
'FnyzTdStr C
TdStrLines A
'TdStrSk C
End Sub

Private Sub Z()
End Sub

