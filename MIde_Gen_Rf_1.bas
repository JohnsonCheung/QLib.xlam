Attribute VB_Name = "MIde_Gen_Rf_1"
Option Explicit
Sub AddRfzPj(DistPj As VBProject)
Dim RfLin, RfLy$()
RfLy = FtLy(RfSrcFfnzDistPj(DistPj))
For Each RfLin In Itr(RfLy)
    AddRf DistPj, RfLin
Next
End Sub

Sub AddRf(A As VBProject, RfLin)
If RfLin = "" Then Exit Sub
Dim F$: F = RfFfn(RfLin)
If HasRfFfn(A, F) Then Exit Sub
A.References.AddFromFile F
End Sub

Function RfFfn$(RfLin)
Dim P%: P = InStr(Replace(RfLin, " ", "-", Count:=3), " ")
RfFfn = Mid(RfLin, P + 1)
End Function

Function HasRfFfn(A As VBProject, RfFfn) As Boolean
Dim R As Vbide.Reference
For Each R In A.References
    If R.FullPath = RfFfn Then HasRfFfn = True: Exit Function
Next
End Function

Function RfSrcFfn$(A As VBProject)
RfSrcFfn = Srcp(A) & "Rf.txt"
End Function

Function RfSrcFfnzDistPj$(DistPj As VBProject)
RfSrcFfnzDistPj = SrcpzDistPj(DistPj) & "Rf.txt"
End Function
Function RfSrcPj() As String()
RfSrcPj = RfSrczPj(CurPj)
End Function
Function RfSrczPj(A As VBProject) As String()
Dim R As Vbide.Reference
For Each R In A.References
    PushI RfSrczPj, RfLin(R)
Next
End Function

Function RfLin$(A As Vbide.Reference)
With A
RfLin = JnSpc(Av(.Name, .Guid, .Major, .Minor, .FullPath))
End With
End Function

