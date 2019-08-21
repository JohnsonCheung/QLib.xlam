Attribute VB_Name = "AAA"
Sub Lis_InvdtMod()
Dim C As VBComponent: For Each C In CPj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        If SubStrCnt(C.Name, "_") = 0 Then Debug.Print C.Name
    End If
Next
End Sub
Sub RenModz()

End Sub
Sub Lis_DupModNm()
Dim Dy()
Dim Cmp As VBComponent: For Each Cmp In CPj.VBComponents
    If Cmp.Type = vbext_ct_StdModule Then
        PushI Dy, Array(AftRev(Cmp.Name, "_"), Cmp.Name)
    End If
Next

Dmp FmtDy(SrtDy(DywDupzC(Dy, 0)), Fmt:=EiSSFmt)
'Stop
End Sub

