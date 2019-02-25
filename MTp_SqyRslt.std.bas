Attribute VB_Name = "MTp_SqyRslt"
Option Explicit
Type SqyRslt: Er() As String: Sqy() As String: End Type
Enum eSqBlkTy
    eErBlk
    ePmBlk
    eSwBlk
    eSqBlk
    eRmBlk
End Enum
Const SqBlkTyNN$ = "ER PM SW SQ RM"
Public SampSqt As New SampSqt
Function SqyRsltzEr(Sqy$(), Er$()) As SqyRslt
SqyRsltzEr.Er = Er
SqyRsltzEr.Sqy = Sqy
End Function

Function SqyRsltzSqTp(SqTp$) As SqyRslt
Dim B() As Blk:            B = BlkAy(SqTp)
Dim PmR As PmRslt:       PmR = PmRsltzLnxAy(LnxAyzBlk(B, "PM"))
Dim Pm As Dictionary: Set Pm = PmR.Pm
Dim SwR As SwRslt:       SwR = SwRsltzLnxAy(LnxAyzBlk(B, "SW"), Pm)
Dim SqR As SqyRslt:      SqR = SqyRsltzGpAy(GpAyzBlkTy(B, "SQ"), Pm, SwR.StmtSw, SwR.FldSw)
Dim Er$():                Er = AyAddAp(ErzBlkAy(B), PmR.Er, SwR.Er, SqR.Er)
                SqyRsltzSqTp = SqyRsltzEr(SqR.Sqy, Er)
End Function

Private Function LnxAyzBlk(A() As Blk, BlkTy$) As Lnx()
Dim J%
For J = 0 To UB(A)
    If A(J).BlkTy = BlkTy Then LnxAyzBlk = A(J).Gp.LnxAy: Exit Function
Next
End Function



