Attribute VB_Name = "QDao_Db_DbInf"
Option Explicit
Private Const CMod$ = "MDao_Db_DbInf."
Private Const Asm$ = "QDao"

Sub BrwDbInf(A As Database)
BrwDs DbInf(A), 2000, BrkColVbl:="TblFld Tbl"
End Sub

Function DbInf(A As Database) As Ds
Dim O As Ds, T$()
T = Tny(A)
AddDt O, InfDtOfLnk(A, T)
AddDt O, InfDtOfTbl(A, T)
AddDt O, InfDtOfTblF(A, T)
AddDt O, InfDtOfPrp(A)
AddDt O, InfDtOfFld(A, T)
O.DsNm = A.Name
DbInf = O
End Function

Private Sub Z_BrwDbInf()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
Dim D As Database
BrwDbInf D
End Sub

Private Sub Z_InfDtOfTbl()
Dim D As Database
Stop
DmpDt InfDtOfTbl(D, Tny(D))
End Sub

Private Function InfDtOfTbl(A As Database, Tny$()) As Dt
Dim T$, Dry(), I
For Each I In Tny
    T = I
    Push Dry, Array(T, NReczT(A, T), TblDes(A, T), StruzT(A, T))
Next
InfDtOfTbl = DtzFF("DbTbl", "Tbl RecCnt Des Stru", Dry)
End Function

Private Function InfDtOfLnk(A As Database, Tny$()) As Dt
Dim T, Dry(), C$
For Each T In Tni(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Dim O As Dt
InfDtOfLnk = DtzFF("DbLnk", "Tbl Connect", Dry)
End Function

Private Function InfDtOfPrp(A As Database) As Dt
Dim Dry()
InfDtOfPrp = DtzFF("DbPrp", "Prp Ty Val", Dry)
End Function
Private Function InfDtOfFld(A As Database, Tny$()) As Dt
Dim Dry(), T
For Each T In Tni(A)
Next
InfDtOfFld = DtzFF("DbFld", "Tbl Fld Pk Ty Si Dft Req Des", Dry)
End Function

Private Function InfDtOfTblF(D As Database, Tny$()) As Dt
Dim Dry()
Dim T$, I
For Each I In Tni(D)
    T = I
    PushIAy Dry, InfDryOfTblF(D, T)
Next
InfDtOfTblF = DtzFF("TblFld", "Tbl Seq Fld Ty Si ", Dry)
End Function

Private Function InfDryOfTblF(D As Database, T) As Variant()
Dim F$, Seq%, I
For Each I In Fny(D, T)
    F = I
    Seq = Seq + 1
    Push InfDryOfTblF, InfDrOfTblF(T, Seq, FdzTF(D, T, F))
Next
End Function

Private Function InfDrOfTblF(T, Seq%, F As Dao.Field2) As Variant()
InfDrOfTblF = Array(T, Seq, F.Name, DtaTy(F.Type))
End Function
Private Sub ZZ()
MDao_Z_Db_DbInf:
End Sub

Private Function InfDtOfLnkLy(A As Database) As String()
Dim T$, I
For Each I In Tny(A)
    T = I
    PushNonBlank InfDtOfLnkLy, CnStrzT(A, T)
Next
End Function
