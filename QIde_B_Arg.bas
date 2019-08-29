Attribute VB_Name = "QIde_B_Arg"
Option Explicit
Option Compare Text

Function ArgAy(MthLin) As String()
ArgAy = SplitCommaSpc(MthPm(MthLin))
End Function

Function MthPm$(MthLin)
MthPm = BetBkt(MthLin)
End Function

Function ArgAyzMthPmAy(MthPmAy$()) As String()
Dim MthPm: For Each MthPm In Itr(MthPmAy)
    PushIAy ArgAyzMthPmAy, SplitCommaSpc(MthPm)
Next
End Function

Function MthPmAy(MthLinAy$()) As String()
Dim MthLin: For Each MthLin In Itr(MthLinAy)
    PushI MthPmAy, BetBkt(MthLin)
Next
End Function

Function ArgAyzMthLinAy(MthLinAy$()) As String()
Dim MthLin: For Each MthLin In Itr(MthLinAy)
    PushIAy ArgAyzMthLinAy, ArgAy(MthLin)
Next
End Function

Function ArgAyP() As String()
ArgAyP = ArgAyzP(CPj)
End Function

Function ArgAyzP(P As VBProject) As String()
ArgAyzP = ArgAyzMthLinAy(MthLinAyzP(P))
End Function

Function ArgAyzPmAy(PmAy$()) As String()
Dim Pm, Arg
For Each Pm In Itr(PmAy)
    For Each Arg In Itr(SplitCommaSpc(Pm))
        PushI ArgAyzPmAy, Arg
    Next
Next
End Function

Function NArg(MthLin) As Byte
NArg = Si(SplitComma(BetBkt(MthLin)))
End Function


'
