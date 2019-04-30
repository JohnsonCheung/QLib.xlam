Attribute VB_Name = "MVb_Fs_Tmp"
Option Explicit

Function TmpCmd$(Optional Fdr$, Optional Fnn$)
TmpCmd = TmpFfn(".cmd", Fdr, Fnn)
End Function

Function TmpFcsv$(Optional Fdr$, Optional Fnn$)
TmpFcsv = TmpFfn(".csv", Fdr, Fnn)
End Function

Function TmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
Fnn = IIf(Fnn0 = "", TmpNm, Fnn0)
TmpFfn = TmpFdr(Fdr) & Fnn & Ext
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Function TmpFx$(Optional Fdr$, Optional Fnn$)
TmpFx = TmpFfn(".xlsx", Fdr, Fnn)
End Function

Function TmpFxm$(Optional Fdr$, Optional Fnn0$)
TmpFxm = TmpFfn(".xlsm", Fdr, Fnn0)
End Function
Property Get TmpRoot$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpRoot = X
End Property

Property Get TmpHom$()
Static X$
If X = "" Then X = TmpRoot & "JC\": EnsPth X
TmpHom = X
End Property

Sub BrwTmpHom()
BrwPth TmpHom
End Sub

Function TmpNmzWithSfx$(Optional Pfx$ = "N")
Static X&
TmpNmzWithSfx = TmpNm(Pfx) & "_" & X
End Function

Function TmpNm$(Optional Pfx$ = "N")
TmpNm = Pfx & Format(Now(), "YYYYMMDD_HHMMSS")
End Function

Function TmpFdr$(Fdr$)
TmpFdr = AddFdrEns(TmpHom, Fdr)
End Function

Property Get TmpPth$()
TmpPth = EnsPth(TmpHom & TmpNm & "\")
End Property

Sub TmpBrwPth()
BrwPth TmpPth
End Sub

Property Get TmpPthFix$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthFix = X
End Property

Property Get TmpPthHom$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthHom = X
End Property
