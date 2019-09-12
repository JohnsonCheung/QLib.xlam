Attribute VB_Name = "MxTouchMd"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxTouchMd."
Sub TouchMdM()
TouchMdzM CMd
End Sub
Private Sub TouchMdzM(M As CodeModule)
If IsMdCached(M) Then Exit Sub
End Sub

Function IsMdCachedM() As Boolean
IsMdCachedM = IsMdCached(CMd)
End Function

Function IsMdCached(M As CodeModule) As Boolean
Dim F$: F = SrcFfn(M.Parent)
If NoFfn(F) Then Exit Function
Dim S$(): S = LyzFt(SrcFfn(M.Parent))
Dim N&: N = M.CountOfLines
Debug.Print N
Debug.Print Si(S)
If N <> Si(S) Then Exit Function
If N = 0 Then IsMdCached = True
If M.Lines(1, N) <> JnCrLf(S) Then Exit Function
IsMdCached = True
End Function
