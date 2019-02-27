Attribute VB_Name = "MVb_Ay_Map_Tak"
Option Explicit

Function AyTakBefDD(A) As String()
'AyTakBefDD = SyzAyMap(A, "TakBefDD")
End Function

Function AyTakAftDot(A) As String()
Dim I
For Each I In Itr(A)
    Push AyTakAftDot, TakAftDot(A)
Next
End Function

Function AyTakAft(A, Sep$) As String()
Dim I
For Each I In Itr(A)
    PushI AyTakAft, TakAft(I, Sep)
Next
End Function

Function AyTakBef(A, Sep$) As String()
Dim I
For Each I In Itr(A)
    PushI AyTakBef, TakBef(I, Sep)
Next
End Function

Function AyTakBefDot(A) As String()
Dim X
For Each X In Itr(A)
    PushI AyTakBefDot, TakBefDot(X)
Next
End Function

Function AyTakBefOrAll(A, Sep$) As String()
Dim I
For Each I In Itr(A)
    Push AyTakBefOrAll, TakBefOrAll(I, Sep)
Next
End Function

Function AyTakT1(A) As String()
Dim L
For Each L In Itr(A)
    PushI AyTakT1, T1(L)
Next
End Function

Function AyTakT2(A) As String()
'AyTakT2 = SyzAyMap(A, "T2")
End Function

Function AyTakT3(A) As String()
'AyTakT3 = SyzAyMap(A, "T3")
End Function

Function AyTakBetBkt(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushI AyTakBetBkt, TakBetBkt(I)
Next
End Function

