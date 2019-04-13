Attribute VB_Name = "MVb_Ay_Map_Tak"
Option Explicit

Function AyTakBefDD(A) As String()
'AyTakBefDD = SyzAyMap(A, "StrBefDD")
End Function

Function AyTakAftDot(A) As String()
Dim I
For Each I In Itr(A)
    Push AyTakAftDot, StrAftDot(A)
Next
End Function

Function AyTakAft(A, Sep$) As String()
Dim I
For Each I In Itr(A)
    PushI AyTakAft, StrAft(I, Sep)
Next
End Function

Function AyTakAftOrAll(A, Sep$) As String()
Dim I
For Each I In Itr(A)
    PushI AyTakAftOrAll, StrAftOrAll(I, Sep)
Next
End Function

Function AyTakBef(A, Sep$) As String()
Dim I
For Each I In Itr(A)
    PushI AyTakBef, StrBef(I, Sep)
Next
End Function

Function AyTakBefDot(A) As String()
Dim X
For Each X In Itr(A)
    PushI AyTakBefDot, StrBefDot(X)
Next
End Function

Function AyTakBefOrAll(A, Sep$) As String()
Dim I
For Each I In Itr(A)
    Push AyTakBefOrAll, StrBefOrAll(I, Sep)
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
    PushI AyTakBetBkt, BetBkt(I)
Next
End Function

