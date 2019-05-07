Attribute VB_Name = "QIde_Mth_Fb_Tmp_Tbl"
Option Explicit
Private Const CMod$ = "MIde_Mth_Fb_Tmp_Tbl."
Private Const Asm$ = "QIde"
'Option Explicit
'Const CMod$ = "MIde_Mth_Fb_Tmp_Tbl."
'Const IMthMch$ = "MthMch"
'Const OMthMd$ = "MthMd"
'Const YFny$ = "MthNm MdNm MchTyStr MchStr"
'
'Property Get AAAModDic() As Dictionary
'' Return a dic of Key=MdNm and Val=LineszMd from Md_MthNmDic("AAAMod")
'' Use #MthMd : MthNm MdNm
'Dim O As Dictionary
''    Set O = JnStrDic_DbTwoColSql(W, "Select MdNm,MthNm from [#MthMd]")
'    Dim K, MthNy$(), MthNmDic As Dictionary
'    'Set MthNmDic = Md_MthNmDic(MdzPj(Pj("QFinal"), "AAAMod"))
'    For Each K In O.Keys
'        If IsNull(K) Then Stop
'        MthNy = CvSy(O(K)) ' The value of the dic is MthNy
'        O(K) = ValzDicIfKyJn(MthNmDic, MthNy) ' return a LineszMd from MthNmDic using MthNy to look MthNmDic
'    Next
'Set AAAModDic = O
'End Property
'
'Sub BrwTmpMthMd()
'BrwQ SqlSel_FF_Fm("MdNm MthNm", "#MthMd")
'End Sub
'
'Sub BrwTmpMthNy()
'BrwTT "#MthNy"
'End Sub
'
'Private Property Get MthNy() As String()
'Stop
''MthNy = SyzTblCol("#MthNy", "MthNy")
'End Property
'
'Sub RfhTmpMthMd()
'RfhTmpMthNy
'Const T$ = "#MthMd"
''DrpTblD W, T
''W.Execute "Create Table [#MthMd] (MthNm Text Not Null, MdNm Text(31),MthMchStr Text)"
''W.Execute SqlCrtSk(T, "MthNm")
''AyIns_Dbt MthNy, W, T ' MthNy is from #MthNy
'Stop '
'XUpd
'End Sub
'
'Sub RfhTmpMthNy()
'Const T$ = "#MthNy"
''DrpTblD W, T
''W.Execute "Create Table [#MthNy] (MthNm Text)"
''W.Execute SqlCrtSk(T, "MthNm")
''AyIns_Dbt MthNyzMd(Md("AAAMod")), W, T
'Stop '
'End Sub
'
'Private Sub WDrp_TmpMthNy()
''DrpTblD W, "#MthNy"
'End Sub
'
'Private Property Get WMchDic() As Dictionary
'Static X As Dictionary
''If IsNothing(X) Then Set X = JnStrDic_DbTwoColSql(W, "Select MthMchStr,ToMdNm from MthMch order by Seq,MthMchStr")
'Set WMchDic = X
'End Property
'
'Private Function WMthNyzMd(A As CodeModule) As String()
''WMthNy_EnsCache A
''WMthNyzMd = ColSyD(W, "#MthNy", "MthNm")
'End Function
'
'Private Sub WMthNy_EnsCache(A As CodeModule)
'Const T$ = "#MthNy"
''If HasDbt(W, T) Then Exit Sub
'W.Execute "Create Table [#MthNy] (MthNm Text(255) Not Null)"
''W.Execute SqlCrtSk(T, "MthNm")
''AyIns_Dbt MthNyzMd(A), W, T
'End Sub
'
'Private Function XDr(MthNm, XDic As Dictionary) As Variant()
''With StrDicMch(MthNm, XDic)
''    If .Patn = "" Then
''        XDr = Array(MthNm, "", "AAAMod")
''    Else
''        XDr = Array(MthNm, .Patn, .Rslt)
''    End If
''End With
'End Function
'
'Sub XUpd()
'Const CSub$ = CMod & "XUpd"
'Dim Rs As DAO.Recordset, XDic As Dictionary
''Set XDic = JnStrDic_DbTwoColSql(W, "Select MthMchStr,ToMdNm from MthMch order by Seq Desc,Ty")
''Set Rs = RsDbq(W, "Select MthNm,MthMchStr,MdNm from [#MthMd] where IIf(IsNull(MthMchStr),'',MthMchStr)=''")
'While Not Rs.EOF
'    'Dr_Upd_Rs XDr(Rs.Fields("MthNm").Value, XDic), Rs
'    Rs.MoveNext
'Wend
'Dim A%, B%, C%
''A = NRecDT(W, "#MthNy")
''B = NRecDT(W, "#MthMd")
''C = NRecDT(W, "#MthMd", "MdNm='AAAMod'")
''C = ValzQ(W, "Select count(*) from [#MthMd] where MdNm='AAAMod'")
'Debug.Print CSub, "A: #MthNy-Cnt "; A
'Debug.Print CSub, "B: #MthMd-Cnt "; B
'Debug.Print CSub, "C: #MthMd-Wh-MdNm=AAAMod-Cnt "; C
''BrwDbq MthDb, "select * from [#MthMd] where MthMchStr='' order by MthNm"
'End Sub
'
'Private Property Get ZZMd() As CodeModule
'Set ZZMd = Md("AAAMod")
'End Property
'
'Private Property Get ZZMthNy() As String()
'ZZMthNy = WMthNyzMd(ZZMd)
'End Property
'
'Private Sub Z_ZZMthNy()
'Brw ZZMthNy
'End Sub
'
'Private Sub ZZ()
'RfhTmpMthMd
'RfhTmpMthNy
'XUpd
'End Sub
'
'Private Sub Z()
'Z_ZZMthNy
'End Sub
