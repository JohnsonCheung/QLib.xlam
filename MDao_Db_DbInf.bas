Attribute VB_Name = "MDao_Db_DbInf"
Option Explicit
Sub BrwDbInf(A As Database)
DbInfDs(A).Brw 2000, DtBrkColDicVbl:="TblFld Tbl"
End Sub

Function DbInfDs(A As Database) As Ds
Dim O As Ds, T$()
T = Tny(A)
DsAddDt O, XLnk(A, T)
DsAddDt O, XTbl(A, T)
DsAddDt O, XTblF(A, T)
DsAddDt O, XPrp(A)
DsAddDt O, XFld(A, T)
O.DsNm = A.Name
DbInfDs = O
End Function

Private Sub Z_BrwDbInf()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
Dim D As Database
BrwDbInf D
End Sub

Private Sub Z_XTbl()
Dim D As Database
Stop
DmpDt XTbl(D, Tny(D))
End Sub

Private Function XTbl(A As Database, Tny$()) As Dt
Dim T, Dry()
For Each T In Tny
    Push Dry, Array(T, NReczT(A, T), TblDes(A, T), StruzT(A, T))
Next
Set XTbl = Dt("DbTbl", "Tbl RecCnt Des Stru", Dry)
End Function

Private Function XLnk(A As Database, Tny$()) As Dt
Dim T, Dry(), C$
For Each T In Tni(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Dim O As Dt
Set XLnk = Dt("DbLnk", "Tbl Connect", Dry)
End Function

Private Function XPrp(A As Database) As Dt
Dim Dry()
Set XPrp = Dt("DbPrp", "Prp Ty Val", Dry)
End Function
Private Function XFld(A As Database, Tny$()) As Dt
Dim Dry(), T
For Each T In Tni(A)
Next
Set XFld = Dt("DbFld", "Tbl Fld Pk Ty Si Dft Req Des", Dry)
End Function

Private Function XTblF(D As Database, Tny$()) As Dt
Dim Dry()
Dim T
For Each T In Tni(D)
    PushIAy Dry, XTblFDry(D, T)
Next
Set XTblF = Dt("TblFld", "Tbl Seq Fld Ty Si ", Dry)
End Function

Private Function XTblFDry(D As Database, T) As Variant()
Dim F, Seq%
For Each F In Fny(D, T)
    Seq = Seq + 1
    Push XTblFDry, XTblFDr(T, Seq, FdzTF(D, T, F))
Next
End Function
Private Function XTblFDr(T, Seq%, F As Dao.Field2) As Variant()
XTblFDr = Array(T, Seq, F.Name, DtaTy(F.Type))
End Function
Private Sub Z()
MDao_Z_Db_DbInf:
End Sub

Private Function XLnkLy(A As Database) As String()
Dim T
For Each T In Tny(A)
    PushNonBlankStr XLnkLy, CnStrzT(A, T)
Next
End Function

