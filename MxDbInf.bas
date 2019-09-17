Attribute VB_Name = "MxDbInf"
Option Compare Text
Option Explicit
Const CNs$ = "sdfsdf"
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxDbInf."

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

Function DroInfoTblF(T, Seq%, F As DAO.Field2) As Variant()
DroInfoTblF = Array(T, Seq, F.Name, DtaTy(F.Type))
End Function

Function DtoInfoFld(D As Database, Tny$()) As Dt
Dim Dy(), T
For Each T In Tni(D)
Next
DtoInfoFld = DtzFF("DbFld", "Tbl Fld Pk Ty Si Dft Req Des", Dy)
End Function

Function DtoInfoLnk(D As Database, Tny$()) As Dt
Dim T, Dy(), C$
For Each T In Tni(D)
   C = D.TableDefs(T).Connect
   If C <> "" Then Push Dy, Array(T, C)
Next
Dim O As Dt
DtoInfoLnk = DtzFF("DbLnk", "Tbl Connect", Dy)
End Function

Function DtoInfoLnkLy(D As Database) As String()
Dim T$, I
For Each I In Tny(D)
    T = I
    PushNB DtoInfoLnkLy, CnStrzT(D, T)
Next
End Function

Function DtoInfoPrp(D As Database) As Dt
Dim Dy()
DtoInfoPrp = DtzFF("DbPrp", "Prp Ty Val", Dy)
End Function

Function DtoInfoTbl(D As Database, Tny$()) As Dt
Dim T$, Dy(), I
For Each I In Tny
    T = I
    Push Dy, Array(T, NReczT(D, T), TdDes(D, T), StruzT(D, T))
Next
DtoInfoTbl = DtzFF("DbTbl", "Tbl RecCnt Des Stru", Dy)
End Function

Function DtoInfoTblF(D As Database, Tny$()) As Dt
Dim Dy()
Dim T$, I
For Each I In Tni(D)
    T = I
    PushIAy Dy, DyoInfoTblFTblF(D, T)
Next
DtoInfoTblF = DtzFF("TblFld", "Tbl Seq Fld Ty Si ", Dy)
End Function

Function DyoInfoTblFTblF(D As Database, T) As Variant()
Dim F$, Seq%, I
For Each I In Fny(D, T)
    F = I
    Seq = Seq + 1
    Push DyoInfoTblFTblF, DroInfoTblF(T, Seq, FdzTF(D, T, F))
Next
End Function


Sub Z_BrwDbInf()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
BrwDbInf SampDb
End Sub

Sub Z_DtoInfoTbl()
Dim D As Database
Stop
DmpDt DtoInfoTbl(D, Tny(D))
End Sub
