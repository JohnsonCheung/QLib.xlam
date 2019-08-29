Attribute VB_Name = "QXls_F_InspDrs"
Option Explicit
Option Compare Text
Enum EmFixWdt
    EiNotFix = 0
    EiFixWdt = 1
End Enum
Private Type X
Wb As Workbook
IxWs As Worksheet
IxLo As ListObject
End Type
Private X As X
Private Sub Init()
EnsWb
EnsIxWs
End Sub
Sub ClrInsp()
Init
While X.Wb.Sheets.Count > 1
    DltWs X.Wb, 2
Wend
ClrLo X.IxLo
End Sub
Sub InspV(V, Optional N$ = "Var")
Init
Dim Las&
With X.IxLo.ListRows
    .Add
    Las = X.IxLo.ListRows.Count
    Dim S$
    If IsStr(V) Then
        S = "'" & V
    Else
        S = V
    End If
    .Item(Las).Range.Value = SqhzAp(Las, N, Empty, TypeName(V), S, Empty, Empty, Empty)
End With
End Sub

Sub InspDrs(A As Drs, N$, Optional Wdt As EmFixWdt = EiNotFix)
Init
Dim Las&, Wsn$, DrsNo%, R As Range
DrsNo = XNxtDrsNo(N)
With X.IxLo.ListRows
    .Add
    Las = X.IxLo.ListRows.Count
    .Item(Las).Range.Value = SqhzAp(Las, N, DrsNo, "Drs", "Go", NRowzDrs(A), NColzDrs(A), IsSamDrEleCnt(A))
End With
Wsn = N & DrsNo
Set R = DtaRgzWs(AddWszDrs(X.Wb, A, Wsn))
If Wdt = EmFixWdt.EiFixWdt Then
    R.Font.Name = "Courier New"
    R.Font.Size = 9
End If
R.Columns.EntireColumn.AutoFit
AddHypLnk LasRowCell(X.IxLo, "Val"), Wsn
End Sub

Private Function XNxtDrsNo%(DrsNm$)
Dim A As Drs, B As Drs, C As Drs
A = DrszLo(X.IxLo)
B = DwEqSel(A, "Nm", DrsNm, "Nm Drs# ValTy")
C = DwEqE(B, "ValTy", "Drs")
If NoReczDrs(C) Then XNxtDrsNo = 1: Exit Function
XNxtDrsNo = AyMax(IntCol(C, "Drs#")) + 1
End Function

Private Sub EnsIxWs()
Set X.IxWs = FstWs(X.Wb)
If X.IxWs.Name <> "Index" Then
    X.IxWs.Name = "Index"
    Set X.IxLo = CrtLo(X.IxWs, "Seq# Nm Drs# ValTy Val NRow NCol IsSamDrEleCnt")
End If
Set X.IxLo = X.IxWs.ListObjects(1)
End Sub

Private Sub EnsWb()
Set X.Wb = EnsWbzXls(Exl, "Insp")
End Sub

Private Sub Z()
QXls_B_XlsOp_InspDrs:
End Sub

'
