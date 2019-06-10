Attribute VB_Name = "QXls_B_XlsOp_InspDrs"
Option Explicit
Option Compare Text
Enum EmFixWdt
    EiNotFix = 0
    EiFixWdt = 1
End Enum
Private Type X
WB As Workbook
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
While X.WB.Sheets.Count > 1
    DltWs X.WB, 2
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
Set R = DtaRgzWs(AddWszDrs(X.WB, A, Wsn))
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
B = ColEqSel(A, "Nm", DrsNm, "Nm Drs# ValTy")
C = ColEqExlEqCol(B, "ValTy", "Drs")
If NoReczDrs(C) Then XNxtDrsNo = 1: Exit Function
XNxtDrsNo = MaxzAy(IntAyzDrsC(C, "Drs#")) + 1
End Function

Private Sub EnsIxWs()
Set X.IxWs = FstWs(X.WB)
If X.IxWs.Name <> "Index" Then
    X.IxWs.Name = "Index"
    Set X.IxLo = CrtLo(X.IxWs, "Seq# Nm Drs# ValTy Val NRow NCol IsSamDrEleCnt")
End If
Set X.IxLo = X.IxWs.ListObjects(1)
End Sub

Private Sub EnsWb()
Set X.WB = EnsWbzXls(Xls, "Insp")
End Sub

Function EnsWbzXls(Xls As Excel.Application, Wbn$) As Workbook
Dim O As Workbook
Const FxFn$ = "Insp.xlsx"
If HasWbn(Xls, FxFn) Then
    Set O = Xls.Workbooks(FxFn)
Else
    Set O = Xls.Workbooks.Add
    O.SaveAs InstFdr("Insp") & "Insp.xlsx"
End If
Set EnsWbzXls = ShwWb(O)
End Function
Function HasWbn(Xls As Excel.Application, Wbn$) As Boolean
HasWbn = HasItn(Xls.Workbooks, Wbn)
End Function
