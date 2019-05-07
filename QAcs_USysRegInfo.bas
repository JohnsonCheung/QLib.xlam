Attribute VB_Name = "QAcs_USysRegInfo"
Option Explicit
Private Const CMod$ = "MAcs_USysRegInfo."
Private Const Asm$ = "QAcs"
'http://www.utteraccess.com/forum/USysRegInf-table-t353963.html
''able Name = USysRegInf
'Fields: Subkey (text), Type (number), ValName (text), Value (text)
'At least 3 records.
'Subkey in all 3 records = 'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\&NameOfYourAdd-inHere'
'Type in 1st record = '0' then '1' in last 2 records
'ValName is blank in first record, then 'Expression' in second and 'Library' in the thid record.
'Value is blank in first record, then '=NameOfFunctionToOpenFormInYourDatabase()' in the second record and '|ACCDIR\NameOfYourDatabase.mde' in the third record.
'That is the best I can suggest. You may need more records depending on your Add-in. Do not add the single quotes (') in the description of what goes in each record.
'hth,
'Jac"
'SK = 'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\&NameOfYourAdd-inHere
' Rec#  SubKey Type ValName        Value
' 1      SK    0     ""            ""
' 2      SK    1     "Expression"  "={FunNm}()"
' 3      SK    1     "Library"     "|ACCDIDR\{fba}"
Sub CrtTblzUSysRegInf(A As Database)
RunQ A, "Create Table [USysRegInf] (Subky Text,Type Long,ValName Text,Value Text)"
End Sub
Sub EnsTblzUSysRegInf(A As Database)
If HasTbl(A, "USysRegInf") Then CrtTblzUSysRegInf A
End Sub

Sub InstallAddin(A As Database, Fb$, Optional AutoFunNm$ = "AutoExec")
Dim Sk$: Sk = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\&NameOfYourAdd-inHere"
Dim Fba$: Fba = ""
Dim FunNm$
Stop '
RunQQ A, "Insert into [USysRegInf] Values('?',0,'','')"
RunQQ A, "Insert into [USysRegInf] Values('?',1,'Expression','?')", Sk, FunNm
RunQQ A, "Insert into [USysRegInf] Values('?',1,'Library','?')", Sk, Fba
End Sub
