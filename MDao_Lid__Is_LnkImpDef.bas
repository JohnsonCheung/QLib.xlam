Attribute VB_Name = "MDao_Lid__Is_LnkImpDef"
Option Explicit
Function LiFxc(ColNm$, ShtTyLis$, ExtNm$) As LiFxc
Set LiFxc = New LiFxc
LiFxc.Init ColNm, ShtTyLis, ExtNm
End Function

Function LiFxcLnkColStr(LnkColStr) As LiFxc
Dim ColNm$, ShtTyLis$, ExtNm$: Asg2TRst LnkColStr, ColNm, ShtTyLis, ExtNm
Set LiFxcLnkColStr = New LiFxc
LiFxcLnkColStr.Init ColNm, ShtTyLis, RmvSqBkt(RTrim(ExtNm))
End Function

Function LiFxcAy(LnkColVbl$) As LiFxc()
LiFxcAy = LiFxcAyLnkColAy(SplitVbar(LnkColVbl))
End Function

Function LiFxcAyLnkColAy(A$()) As LiFxc()
Dim I
For Each I In Itr(A)
    PushObj LiFxcAyLnkColAy, LiFxcLnkColStr(I)
Next
End Function

Function LiFx(Fxn$, T$, Wsn$, Fxc() As LiFxc, Optional Bexpr) As LiFx
Dim O As New LiFx
Set LiFx = O.Init(Fxn, T, Wsn, Fxc, Bexpr)
End Function

Function LiFxLnkColVbl(Fxn$, T$, Wsn$, LnkColVbl$, Optional Bexpr$) As LiFx
Set LiFxLnkColVbl = LiFx(Fxn, Wsn, T, LiFxcAy(LnkColVbl), Bexpr)
End Function

Function LiFil(FilNm$, Ffn$) As LiFil
Dim O As New LiFil
Set LiFil = O.Init(FilNm, Ffn)
End Function

Function LiFb(Fbn, T, Fset As Aset, Bexpr$) As LiFb
Dim O As New LiFb
Set LiFb = O.Init(Fbn, T, Fset, Bexpr)
End Function

