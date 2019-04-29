Attribute VB_Name = "MVb_Fs_Ffn_Ext"
Option Explicit
Function RplExt$(Ffn$, NewExt)
RplExt = RmvExt(Ffn$) & NewExt
End Function

Sub ThwNotExt(S)
If Not IsExt(S) Then Err.Raise 1, , "Not Ext(" & S & ")"
End Sub

Function IsExt(S) As Boolean
IsExt = FstChr(S) = "."
End Function


