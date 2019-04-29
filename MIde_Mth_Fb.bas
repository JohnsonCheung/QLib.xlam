Attribute VB_Name = "MIde_Mth_Fb"
Option Explicit
Function MthFbInPj$()
MthFbInPj = MthFbzPj(CurPj)
End Function
Function MthFb$()
MthFb = MthFbInPj
End Function
Function MthFbzPj$(A As VBProject)
MthFbzPj = MthPthzPj(A) & Fn(Pjf(A)) & ".MthDb.accdb"
End Function

Function MthPthzPj$(A As VBProject)
Dim F$: F = Pjf(A)
MthPthzPj = AddFdrEns(Pth(F), ".MthDb")
End Function

Function EnsMthFb(MthFb$) As Database
EnsFb MthFb
Dim D As Database
Set EnsMthFb = Db(MthFb)
EnsSchm D, MthSchm
End Function

Function MthDbzPj(A As VBProject) As Database
Dim Fb$: Fb = MthFbzPj(A)
EnsMthFb Fb
Set MthDbzPj = Db(Fb$)
End Function

Property Get MthDbInPj() As Database
Set MthDbInPj = MthDbzPj(CurPj)
End Property

Sub BrwMthFb()
BrwFb MthFb
End Sub

Private Property Get MthSchm() As String()
Erase XX
X "Fld Nm  Md Pj"
X "Fld T50 MchStr"
X "Fld T10 MthPfx"
X "Fld Txt Pjf Prm Ret LinRmk"
X "Fld T3  Ty Mdy"
X "Fld T4  MdTy"
X "Fld Lng Lno"
X "Fld Mem Lines TopRmk"
X "Tbl Pj  *Id Pjf | PjNm PjDte"
X "Tbl Md  *Id PjId MdNm | MdTy"
X "Tbl Mth *Id MdId MthNm ShtTy | ShtMdy Prm Ret LinRmk TopRmk Lines Lno"
MthSchm = XX
Erase XX
End Property

