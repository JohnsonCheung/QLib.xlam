Attribute VB_Name = "QVb_Stop_Xls"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Stop_Xls."
Private Const Asm$ = "QVb"
Declare Function GetCurrentProcessId& Lib "Kernel32.dll" ()
'Declare Function GetProcessId& Lib "Kernel32.dll" (ProcessHandle&)
'Const Ps1Str$ = "function Get-ExcelProcessId { try { (Get-Process -Name Excel).Id } finally { @() } }" & vbCrLf & _
'"Stop-Process -Id (Get-ExcelProcessId)"

Sub StopXls()
Const Ps1Str$ = "Stop-Process -Id{try{(Get-Process -Name Excel).Id}finally{@()}}.invoke()"
Dim F$: F = TmpHom & "StopXls.ps1"
Static X As Boolean
If Not X Then
    X = True
    If Not HasFfn(F) Then
        WrtStr Ps1Str, F
    End If
End If
'PowerRun F
End Sub

