Attribute VB_Name = "MAcs_USysRegInfo"
Option Explicit
'http://www.utteraccess.com/forum/USysRegInfo-table-t353963.html
''able Name = USysRegInfo
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
Sub CrtTblzUSysRegInfo()
RunQz CDb, "Create Table [USysRegInfo] (Subky Text,Type Long,ValName Text,Value Text)"
End Sub
Sub CrtTblzUSysRegInfoDb(A As Database)
RunQz A, "Create Table [USysRegInfo] (Subky Text,Type Long,ValName Text,Value Text)"
End Sub
Sub EnsTblzUSysRegInfoz(A As Database)
If HasTblz(A, "USysRegInfo") Then CrtTblzUSysRegInfo
End Sub

Sub InstallAddIn(A As Database, Fb$, Optional AutoFunNm$ = "AutoExec")
Dim Sk$: Sk = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\&NameOfYourAdd-inHere"
Dim Fba$: Fba = ""
Dim FunNm$
Stop '
RunQQz A, "Insert into [USysRegInfo] Values('?',0,'','')"
RunQQz A, "Insert into [USysRegInfo] Values('?',1,'Expression','?')", Sk, FunNm
RunQQz A, "Insert into [USysRegInfo] Values('?',1,'Library','?')", Sk, Fba
End Sub
