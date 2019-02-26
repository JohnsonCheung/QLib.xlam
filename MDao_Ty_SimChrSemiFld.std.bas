Attribute VB_Name = "MDao_Ty_SimChrSemiFld"
Option Explicit
Function TdTblShtTySemiFldSsl(T, ShtTySemiFldSsl$) As Dao.TableDef
Dim Ay$(): Ay = TermAy(ShtTySemiFldSsl)
Dim FdAy() As Dao.Field2, I
For Each I In TermAy(ShtTySemiFldSsl)
    PushObj FdAy, FdShtTySemiFld(I)
Next
'Set TdTblShtTySemiFld = NewTdTblFdAy(T, FdAy)
End Function

Function FdShtTySemiFld(A) As Dao.Field2
Dim ShtTy$, F$
'AsgBrkColon A, ShtTy, F
Set FdShtTySemiFld = Fd(F, DaoTyzShtTy(ShtTy))
End Function
