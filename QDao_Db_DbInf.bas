Attribute VB_Name = "QDao_Db_DbInf"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Db_DbInf."
Private Const Asm$ = "QDao"

Sub BrwDbInf(D As Database)
BrwDs DbInf(D), 2000, BrkColVbl:="TblFld Tbl"
End Sub

Function DbInf(D As Database) As Ds
Dim O As Ds, T$()
T = Tny(D)
AddDt O, DtoInfoLnk(D, T)
AddDt O, DtoInfoTbl(D, T)
AddDt O, DtoInfoTblF(D, T)
AddDt O, DtoInfoPrp(D)
AddDt O, DtoInfoFld(D, T)
O.DsNm = D.Name
DbInf = O
End Function

Sub Z_BrwDbInf()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
BrwDbInf SampDb
End Sub

Private Sub Z_DtoInfoTbl()
Dim D As Database
Stop
DmpDt DtoInfoTbl(D, Tny(D))
End Sub

Private Function DtoInfoTbl(D As Database, Tny$()) As DT
Dim T$, Dy(), I
For Each I In Tny
    T = I
    Push Dy, Array(T, NReczT(D, T), TblDes(D, T), StruzT(D, T))
Next
DtoInfoTbl = DtzFF("DbTbl", "Tbl RecCnt Des Stru", Dy)
End Function

Private Function DtoInfoLnk(D As Database, Tny$()) As DT
Dim T, Dy(), C$
For Each T In Tni(D)
   C = D.TableDefs(T).Connect
   If C <> "" Then Push Dy, Array(T, C)
Next
Dim O As DT
DtoInfoLnk = DtzFF("DbLnk", "Tbl Connect", Dy)
End Function

Private Function DtoInfoPrp(D As Database) As DT
Dim Dy()
DtoInfoPrp = DtzFF("DbPrp", "Prp Ty Val", Dy)
End Function
Private Function DtoInfoFld(D As Database, Tny$()) As DT
Dim Dy(), T
For Each T In Tni(D)
Next
DtoInfoFld = DtzFF("DbFld", "Tbl Fld Pk Ty Si Dft Req Des", Dy)
End Function

Private Function DtoInfoTblF(D As Database, Tny$()) As DT
Dim Dy()
Dim T$, I
For Each I In Tni(D)
    T = I
    PushIAy Dy, DyoInfoTblFTblF(D, T)
Next
DtoInfoTblF = DtzFF("TblFld", "Tbl Seq Fld Ty Si ", Dy)
End Function

Private Function DyoInfoTblFTblF(D As Database, T) As Variant()
Dim F$, Seq%, I
For Each I In Fny(D, T)
    F = I
    Seq = Seq + 1
    Push DyoInfoTblFTblF, DroInfoTblF(T, Seq, FdzTF(D, T, F))
Next
End Function

Private Function DroInfoTblF(T, Seq%, F As Dao.Field2) As Variant()
DroInfoTblF = Array(T, Seq, F.Name, DtaTy(F.Type))
End Function
Private Sub Z()
MDao_Z_Db_DbInf:
End Sub

Private Function DtoInfoLnkLy(D As Database) As String()
Dim T$, I
For Each I In Tny(D)
    T = I
    PushNB DtoInfoLnkLy, CnStrzT(D, T)
Next
End Function

