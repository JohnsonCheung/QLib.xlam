Attribute VB_Name = "MIde_ConstMth"
Option Explicit
Public Const DocOfQNm$ = "newtype AOptDotNm.  "
Public Const DocOfAOptDotNm$ = "type Nm | ADotNm."

Sub EdtConst(ConstQNm$)
EdtStr ConstVal(ConstQNm), FtzConstQNm(ConstQNm)
End Sub

Sub UpdConst(ConstQNm$, Optional IsPub As Boolean)
With MdMth(ConstQNm)
    RplMth .Md, .MthNm, ConstPrpLines(ConstQNm, IsPub)
End With
End Sub

Private Property Get Z_CrtSchm1() As String()
Erase XX
X "Tbl A *Id | *Nm     | *Dte AATy Loc Expr Rmk"
X "Tbl B *Id | AId *Nm | *Dte"
X "Fld Txt AATy"
X "Fld Loc Loc"
X "Fld Expr Expr"
X "Fld Mem Rmk"
X "Ele Loc Txt Rq Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
X "Ele Expr Txt [Expr=Loc & 'abc']"
X "Des Tbl     A     AA BB "
X "Des Tbl     A     CC DD "
X "Des Fld     ANm   AA BB "
X "Des Tbl.Fld A.ANm TF_Des-AA-BB"
Z_CrtSchm1 = XX
Erase XX
End Property

Private Property Get C_A() As String()
Erase XX
X "lsjdf lskdjf lsdkf"
X "sdfkljsdf"
X "sdf"
X "sdf"
C_A = XX
C_A = XX
Erase XX
End Property
