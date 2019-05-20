Attribute VB_Name = "QIde_Mth_Fb_Gen"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Fb_Gen."
Private Const Asm$ = "QIde"

Sub CrtTblMth()
Dim W As Database
'WDrp "DistMth #A #B"
Q = "Select Distinct Nm,Count(*) as LinesIdCnt Into DistMth from DistLines group by Nm": W.Execute Q
Q = "Alter Table DistMth Add Column LinesIdLis Text(255), LinesLis Memo, ToMd Text(50)": W.Execute Q
'WtCrt_FldLisTbl "DistLines", "#A", "Nm", "LinesId", " ", True
'WtCrt_FldLisTbl "DistLines", "#B", "Nm", "Lines", vbCrLf & vbCrLf, True
Q = "Update DistMth x inner join [#A] a on x.Nm = a.Nm set x.LinesIdLis = a.LinesIdLis":                W.Execute Q
Q = "Update DistMth x inner join [#B] a on x.Nm = a.Nm set x.LinesLis = a.LinesLis":                    W.Execute Q
Q = "Update DistMth x inner join MthLoc a on x.Nm = a.Nm set x.ToMd = IIf(a.ToMd='','AAMod',a.ToMd)":   W.Execute Q
End Sub

Sub CrtMdDic()
'WSetDb MthDb
'WDrp "MdDic"
'WtCrt_FldLisTbl "DistMth", "MdDic", "ToMd", "LinesLis", vbCrLf & vbCrLf, True, "Lines"
End Sub

Sub UpdMthLoc()
Dim W As Database
'WSetDb MthDb
'WDrp "#A"
Q = "Select x.Nm into [#A] from DistMth x left join MthLoc a on x.Nm=a.Nm where IsNull(a.Nm)": W.Execute Q
Q = "Insert into MthLoc (Nm) Select Nm from [#A]": W.Execute Q
End Sub

Private Sub ZZ()
UpdMthLoc
End Sub
