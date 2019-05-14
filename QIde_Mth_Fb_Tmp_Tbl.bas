Attribute VB_Name = "QIde_Mth_Fb_Tmp_Tbl"
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
'    Dim K, Mthny$(), MthnDic As Dictionary
'    'Set MthnDic = Md_MthnDic(MdzP(Pj("QFinal"), "AAAMod"))
'    For Each K In O.Keys
'        If IsNull(K) Then Stop
'        Mthny = CvSy(O(K)) ' The value of the dic is Mthny
'        O(K) = ValzDicIfKyJn(MthnDic, Mthny) ' return a LineszMd from MthnDic using Mthny to look MthnDic
'    Next
'Set AAAModDic = O
'End Property
'
'Sub BrwTmpMthMd()
'BrwQ SqlSel_FF_Fm("Mdn Mthn", "#MthMd")
'End Sub
'
'Sub BrwTmpMthny()
'BrwTT "#Mthny"
'End Sub
'
'Private Property Get Mthny() As String()
'Stop
''Mthny = SyzTblCol("#Mthny", "Mthny")
'End Property
'
'Sub RfhTmpMthMd()
'RfhTmpMthny
'Const T$ = "#MthMd"
''DrpTblD W, T
''W.Execute "Create Table [#MthMd] (Mthn Text Not Null, Mdn Text(31),MthMchStr Text)"
''W.Execute SqlCrtSk(T, "Mthn")
''AyIns_Dbt Mthny, W, T ' Mthny is from #Mthny
'Stop '
'XUpd
'End Sub
'
'Sub RfhTmpMthny()
'Const T$ = "#Mthny"
''DrpTblD W, T
''W.Execute "Create Table [#Mthny] (Mthn Text)"
''W.Execute SqlCrtSk(T, "Mthn")
''AyIns_Dbt MthnyzMd(Md("AAAMod")), W, T
'Stop '
'End Sub
'
'Private Sub WDrp_TmpMthny()
''DrpTblD W, "#Mthny"
'End Sub
'
'Private Property Get WMchDic() As Dictionary
'Static X As Dictionary
''If IsNothing(X) Then Set X = JnStrDic_DbTwoColSql(W, "Select MthMchStr,ToMdn from MthMch order by Seq,MthMchStr")
'Set WMchDic = X
'End Property
'
'Private Function WMthnyzMd(A As CodeModule) As String()
''WMthny_EnsCache A
''WMthnyzMd = ColSyD(W, "#Mthny", "Mthn")
'End Function
'
'Private Sub WMthny_EnsCache(A As CodeModule)
'Const T$ = "#Mthny"
''If HasDbt(W, T) Then Exit Sub
'W.Execute "Create Table [#Mthny] (Mthn Text(255) Not Null)"
''W.Execute SqlCrtSk(T, "Mthn")
''AyIns_Dbt MthnyzMd(A), W, T
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
''A = NRecDT(W, "#Mthny")
''B = NRecDT(W, "#MthMd")
''C = NRecDT(W, "#MthMd", "Mdn='AAAMod'")
''C = ValzQ(W, "Select count(*) from [#MthMd] where Mdn='AAAMod'")
'Debug.Print CSub, "A: #Mthny-Cnt "; A
'Debug.Print CSub, "B: #MthMd-Cnt "; B
'Debug.Print CSub, "C: #MthMd-Wh-Mdn=AAAMod-Cnt "; C
''BrwDbq MthDb, "select * from [#MthMd] where MthMchStr='' order by Mthn"
'End Sub
'
'Private Function Y_Md() As CodeModule
'Set ZZMd = Md("AAAMod")
'End Property
'
'Private Function Y_Mthny() As String()
'ZZMthny = WMthnyzMd(ZZMd)
'End Property
'
'Private Sub Z_ZZMthny()
'Brw ZZMthny
'End Sub
'
'Private Sub ZZ()
'RfhTmpMthMd
'RfhTmpMthny
'XUpd
'End Sub
'
