Attribute VB_Name = "QIde_Mth_Fb_Tmp_Tbl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Fb_Tmp_Tbl."
Private Const Asm$ = "QIde"
'Option Explicit
'Const CMod$ = "MIde_Mth_Fb_Tmp_Tbl."
'Const IMthMch$ = "MthMch"
'Const OMthMd$ = "MthMd"
'Const YFny$ = "Mthn Mdn MchTyStr MchStr"
'
'Property Get AAAModDic() As Dictionary
'' Return a dic of Key=Mdn and Val=LineszMd from Md_MthnDic("AAAMod")
'' Use #MthMd : Mthn Mdn
'Dim O As Dictionary
''    Set O = JnStrDic_DbTwoColSql(W, "Select Mdn,Mthn from [#MthMd]")
'    Dim K, MthNy$(), MthnDic As Dictionary
'    'Set MthnDic = Md_MthnDic(MdzP(Pj("QFinal"), "AAAMod"))
'    For Each K In O.Keys
'        If IsNull(K) Then Stop
'        MthNy = CvSy(O(K)) ' The value of the dic is MthNy
'        O(K) = ValzDicIfKyJn(MthnDic, MthNy) ' return a LineszMd from MthnDic using MthNy to look MthnDic
'    Next
'Set AAAModDic = O
'End Property
'
'Sub BrwTmpMthMd()
'BrwQ SqlSel_FF_Fm("Mdn Mthn", "#MthMd")
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
''W.Execute "Create Table [#MthMd] (Mthn Text Not Null, Mdn Text(31),MthMchStr Text)"
''W.Execute SqlCrtSk(T, "Mthn")
''AyIns_Dbt MthNy, W, T ' MthNy is from #MthNy
'Stop '
'XUpd
'End Sub
'
'Sub RfhTmpMthNy()
'Const T$ = "#MthNy"
''DrpTblD W, T
''W.Execute "Create Table [#MthNy] (Mthn Text)"
''W.Execute SqlCrtSk(T, "Mthn")
''AyIns_Dbt MthNyzM(Md("AAAMod")), W, T
'Stop '
'End Sub
'
'Private Sub WDrp_TmpMthNy()
''DrpTblD W, "#MthNy"
'End Sub
'
'Private Property Get WMchDic() As Dictionary
'Static X As Dictionary
''If IsNothing(X) Then Set X = JnStrDic_DbTwoColSql(W, "Select MthMchStr,ToMdn from MthMch order by Seq,MthMchStr")
'Set WMchDic = X
'End Property
'
'Private Function WMthNyzM(M As CodeModule) As String()
''WMthNy_EnsCache A
''WMthNyzM = ColSyD(W, "#MthNy", "Mthn")
'End Function
'
'Private Sub WMthNy_EnsCache(M As CodeModule)
'Const T$ = "#MthNy"
''If HasDbt(W, T) Then Exit Sub
'W.Execute "Create Table [#MthNy] (Mthn Text(255) Not Null)"
''W.Execute SqlCrtSk(T, "Mthn")
''AyIns_Dbt MthNyzM(A), W, T
'End Sub
'
'Private Function XDr(Mthn, XDic As Dictionary) As Variant()
''With StrDicMch(Mthn, XDic)
''    If .Patn = "" Then
''        XDr = Array(Mthn, "", "AAAMod")
''    Else
''        XDr = Array(Mthn, .Patn, .Rslt)
''    End If
''End With
'End Function
'
'Sub XUpd()
'Const CSub$ = CMod & "XUpd"
'Dim Rs As DAO.Recordset, XDic As Dictionary
''Set XDic = JnStrDic_DbTwoColSql(W, "Select MthMchStr,ToMdn from MthMch order by Seq Desc,Ty")
''Set Rs = RsDbq(W, "Select Mthn,MthMchStr,Mdn from [#MthMd] where IIf(IsNull(MthMchStr),'',MthMchStr)=''")
'While Not Rs.EOF
'    'Dr_Upd_Rs XDr(Rs.Fields("Mthn").Value, XDic), Rs
'    Rs.MoveNext
'Wend
'Dim A%, B%, C%
''A = NRecDT(W, "#MthNy")
''B = NRecDT(W, "#MthMd")
''C = NRecDT(W, "#MthMd", "Mdn='AAAMod'")
''C = ValzQ(W, "Select count(*) from [#MthMd] where Mdn='AAAMod'")
'Debug.Print CSub, "A: #MthNy-Cnt "; A
'Debug.Print CSub, "B: #MthMd-Cnt "; B
'Debug.Print CSub, "C: #MthMd-Wh-Mdn=AAAMod-Cnt "; C
''BrwDbq MthDb, "select * from [#MthMd] where MthMchStr='' order by Mthn"
'End Sub
'
'Private Function Y_Md() As CodeModule
'Set ZZMd = Md("AAAMod")
'End Property
'
'Private Function Y_MthNy() As String()
'ZZMthNy = WMthNyzM(ZZMd)
'End Property
'
'Private Sub Z_ZZMthNy()
'Brw ZZMthNy
'End Sub
'
'Private Sub Z()
'RfhTmpMthMd
'RfhTmpMthNy
'XUpd
'End Sub
'
