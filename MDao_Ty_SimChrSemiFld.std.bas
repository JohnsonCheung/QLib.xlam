Attribute VB_Name = "MDao_Ty_SimChrSemiFld"
Option Explicit
Function TdTblShtTySemiFldSsl(T, ShtTySemiFldSsl$) As DAO.TableDef
Dim Ay$(): Ay = TermAy(ShtTySemiFldSsl)
Dim FdAy() As DAO.Field2, I
For Each I In TermAy(ShtTySemiFldSsl)
    PushObj FdAy, FdShtTySemiFld(I)
Next
'Set TdTblShtTySemiFld = NewTdTblFdAy(T, FdAy)
End Function

Function FdShtTySemiFld(A) As DAO.Field2
Dim ShtTy$, F$
'AsgBrkColon A, ShtTy, F
Set FdShtTySemiFld = NewFd(F, DaoTyzShtTy(ShtTy))
End Function
