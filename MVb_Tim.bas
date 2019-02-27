Attribute VB_Name = "MVb_Tim"
Option Explicit
Private M$, Beg As Date
Public NoStamp As Boolean
Sub TimBeg(Optional Msg$ = "Time")
If M <> "" Then TimEnd
M = Msg
Beg = Now
End Sub
Sub TimEnd(Optional Halt As Boolean)
Debug.Print M & " " & DateDiff("S", Beg, Now) & "(s)"
If Halt Then Stop
End Sub
Sub Stamp(S$)
If Not NoStamp Then Debug.Print NowStr; " "; S
End Sub

