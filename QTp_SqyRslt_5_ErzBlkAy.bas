Attribute VB_Name = "QTp_SqyRslt_5_ErzBlkAy"
Option Explicit
Private Const CMod$ = "MTp_SqyRslt_5_ErzBlkAy."
Private Const Asm$ = "QTp"

Function ErzBlkAy(A() As Blk) As String()
Dim GpAy() As Gp: GpAy = GpAyzBlkTy(A, "Er")
If Si(GpAy) > 0 Then
    PushIAy ErzBlkAy, ErzGpAy(GpAy, "Unexpected Blk, valid block is PM SW SQ RM")
End If
End Function
Private Function CvGpAy(A) As Gp()
CvGpAy = A
End Function
Private Function ErzExcessPmBlk(A() As Blk) As String()
Dim GpAy() As Gp: GpAy = GpAyzBlkTy(A, "PM")
If Si(GpAy) > 1 Then
    PushIAy ErzExcessPmBlk, ErzGpAy(CvGpAy(AyeFstEle(GpAy)), "Excess Pm block, they are ignored")
End If
End Function

Private Function ErzExcessSwBlk(A() As Blk) As String()
Dim GpAy() As Gp: GpAy = GpAyzBlkTy(A, "SW")
If Si(GpAy) > 1 Then
    PushIAy ErzExcessSwBlk, ErzGpAy(CvGpAy(AyeFstEle(GpAy)), "Excess Sw block, they are ignored")
End If
End Function

Function ErzGpAy(GpAy() As Gp, Msg$) As String()

End Function
