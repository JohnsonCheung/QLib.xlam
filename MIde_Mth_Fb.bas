Attribute VB_Name = "MIde_Mth_Fb"
Option Explicit
Public Const MthLocFx$ = "C:\Users\User\Desktop\Vba-Lib-1\MthLoc.xlsx"
Public Const MthFb$ = "C:\Users\User\Desktop\Vba-Lib-1\Mth.accdb"
Private Const WrkFb$ = "C:\Users\User\Desktop\Vba-Lib-1\MthWrk.accdb"

Sub MthFbEns()
EnsMthFb MthFb
End Sub

Function EnsMthFb(MthFb$) As Database
EnsFb MthFb
Dim Db As Database
Set EnsMthFb = Db(MthFb)
'EnsSchmz Db, MthSchm
End Function

Property Get MthDb() As Database
Set MthDb = Db(MthFb)
End Property

Sub BrwMthFb()
BrwFb MthFb
End Sub

Private Property Get MthSchm$()
Const A_1$ = "Fld Nm  Md Pj" & _
vbCrLf & "Fld T50 MchStr" & _
vbCrLf & "Fld T10 MthPfx" & _
vbCrLf & "Fld Txt Pjf Prm Ret LinRmk" & _
vbCrLf & "Fld T3  Ty Mdy" & _
vbCrLf & "Fld T4  MdTy" & _
vbCrLf & "Fld Lng Lno" & _
vbCrLf & "Fld Mem Lines TopRmk" & _
vbCrLf & "Tbl Pj  *Id Pjf | PjNm PjDte" & _
vbCrLf & "Tbl Md  *Id PjId MdNm | MdTy" & _
vbCrLf & "Tbl Mth *Id MdId MthNm ShtTy | ShtMdy Prm Ret LinRmk TopRmk Lines Lno" & _
vbCrLf & ""
MthSchm = A_1
End Property

