Attribute VB_Name = "MIde_Gen_Rf_1"
Option Explicit
Sub AddRfzPj(A As VBProject)
Dim F$: F$ = RfSrcFfn(A)
If Not HasFfn(F) Then Err.Raise 1, , "No RfSrcFfn(" & F & ")"
Dim RfLin
For Each RfLin In Itr(FtLy(F))
    AddRfzRfLin A, RfLin
Next
End Sub
Sub AddRfzRfLin(A As VBProject, RfLin)
Dim Ffn$: Ffn = RfFfn(RfLin)
If HasRfFfn(A, Ffn) Then Exit Sub
A.References.AddFromFile Ffn
End Sub

Function RfFfn$(RfLin)
Dim P%: P = InStr(Replace(RfLin, " ", "-", Count:=3), " ")
RfFfn = Mid(RfLin, P + 1)
End Function

Function HasRfFfn(A As VBProject, RfFfn) As Boolean
Dim R As VBIDE.Reference
For Each R In A.References
    If R.FullPath = RfFfn Then HasRfFfn = True: Exit Function
Next
End Function

Function RfSrcFfn$(A As VBProject)
RfSrcFfn = SrcPth(A) & "Rf.txt"
End Function

Function RfSrc(A As VBProject) As String()
Dim R As VBIDE.Reference
For Each R In A.References
    PushI RfSrc, RfLin(R)
Next
End Function

Function RfLin$(A As VBIDE.Reference)
With A
RfLin = JnSpc(Av(.Name, .Guid, .Major, .Minor, .FullPath))
End With
End Function

