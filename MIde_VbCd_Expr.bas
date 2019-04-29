Attribute VB_Name = "MIde_VbCd_Expr"
Option Explicit
Private Type LinRslt
    ExprLin As String
    OvrFlwTerm As String
    S As String
End Type
Private Type Term
    ExprTerm As String
    S As String
End Type

Function ExprLyzStr(Str, Optional MaxCdLinWdt% = 200) As String()
Dim L, Ay$(): Ay = SplitCrLf(Str)
Dim J&, Fst As Boolean
Erase XX
Fst = True
For Each L In Itr(Ay)
    If Fst Then
        Fst = False
    Else
        X J & ":" & Len(L) & ":" & L
    End If
'    Stop
    J = J + 1
'    PushIAy ExprLyzStr, ExprLyzLin(L, MaxCdLinWdt)
'    Stop
Next
Brw SyAddIxPfx(XX)
Stop
Erase XX
End Function
Private Function ExprLyzLin(Lin, W%) As String()
Dim J&
Dim S$: S = Lin
Dim CurLen&
Dim LasLen&: LasLen = Len(S)
Dim OvrFlwTerm$
While LasLen > 0
    DoEvents
    J = J + 1: If J > 10000 Then ThwLoopingTooMuch CSub
    Stop
    If J > 10 Then Stop
    With ShfLin(S, OvrFlwTerm, W)
        If .ExprLin = "" Then Exit Function
        PushI ExprLyzLin, .ExprLin
        S = .S
        OvrFlwTerm = .OvrFlwTerm
    End With
    CurLen = Len(S)
    If CurLen >= LasLen Then ThwIfNEver CSub, "Str is not shifted by ShfLin"
    LasLen = CurLen
Wend
End Function

Private Function ShfLin(Str$, OvrFlwTerm$, W%) As LinRslt
Dim T$, OExprTermAy$(), TotW&
If OvrFlwTerm <> "" Then
    PushI OExprTermAy, OvrFlwTerm
    TotW = Len(OvrFlwTerm) + 3
End If
Dim S$: S = Str
Dim J&, OStr$, OExprTerm$

X:
ShfLin = LinRslt(ExprLin:=Jn(OExprTermAy, " & "), OvrFlwTerm:=OvrFlwTerm, S:=OStr)
End Function
Private Function Z_ShfTermzPrintable()
Dim S$: S = StrzCurPjf
Dim Las&, Cur&, O$()
Las = Len(S)
While Len(S) > 0
    PushI O, ShfTermzPrintable(S)
    Cur = Len(S)
    If Cur >= Las Then Stop
    Las = Cur
Wend
MsgBox Si(O)
Stop
Brw O
End Function
Private Function ShfTermzPrintable$(OStr$)
If OStr = "" Then Exit Function
Dim IsPrintable As Boolean
Dim J&
IsPrintable = IsAscPrintable(Asc(FstChr(OStr)))
For J = 2 To Len(OStr)
    If IsPrintable <> IsAscPrintable(Asc(Mid(OStr, J, 1))) Then
        ShfTermzPrintable = Left(OStr, J - 1)
        OStr = Mid(OStr, J)
        Exit Function
    End If
Next
ShfTermzPrintable = OStr
OStr = ""
End Function

'Fun=================================================
Private Function LinRslt(ExprLin$, OvrFlwTerm$, S$) As LinRslt
With LinRslt
    .ExprLin = ExprLin
    .OvrFlwTerm = OvrFlwTerm
    .S = S
End With
End Function


Private Function ExprzQuote$(BytAy() As Byte)
Dim O$(), I
For Each I In BytAy
    If I = vbDblQuoteAsc Then PushI O, vbTwoDblQuote Else PushI O, Chr(I)
Next
ExprzQuote = QuoteDbl(Jn(O))
End Function

Private Function ExprzAndChr$(BytAy() As Byte)
Dim O$(), I
For Each I In BytAy
    PushI O, "Chr(" & I & ")"
Next
ExprzAndChr = Jn(O, " & ")
End Function

Private Function Term(ExprTerm$, S$) As Term
With Term
    .ExprTerm = ExprTerm
    .S = S
End With
End Function
Private Sub Z_ExprLyzStr()
Dim S$
GoSub ZZ1
GoSub ZZ2
GoSub T0
GoSub T1
Exit Sub
ZZ2:
    S = StrzCurPjf
    Brw ExprLyzStr(S)
    Return
ZZ1:
    S = StrzCurPjf
    Brw ExprLyzStr(S)
    Return
T0:
    S = "lksdjf lskdf dkf " & Chr(2) & Chr(11) & "ksldfj"
    Ept = Sy("")
    GoTo Tst
T1:
    GoTo Tst
Tst:
    Act = ExprLyzStr(S)
    D Act
    Stop
    C
    Return
End Sub
Private Sub AAA()
Dim J%, A
Erase XX
For J = 0 To 255
    X FmtQQ("If Asc(""?"")<>? Then Debug.Print ?", Chr(J), J, J)
Next
Brw XX
Erase XX
End Sub
Private Sub Z_BrwRepeatedBytes()
BrwRepeatedBytes StrzCurPjf
End Sub

Function AscStr$(S)
Dim J&, O$()
For J = 1 To Len(S)
    PushI O, Asc(Mid(S, J, 1))
Next
AscStr = JnSpc(O)
End Function

Private Sub Z_BrkAyzPrintable1()
Dim T, O$(), J&
For Each T In BrkAyzPrintable(JnCrLf(SrcInPj))
    J = J + 1
    Push O, FmtPrintableStr(T)
Next
Brw SyAddIxPfx(O)
End Sub

Function FmtPrintableStr$(T)
Dim S$: S = PrintableSts(T)
Dim P$: P = S & " " & AlignL(Len(T), 8) & " : "
Select Case S
Case "Prt": FmtPrintableStr = P & T
Case "Non": FmtPrintableStr = P & AscStr(Left(T, 10))
Case "Mix": FmtPrintableStr = P & AscStr(Left(T, 10))
Case Else
    Stop
End Select
End Function
Private Sub Z_BrkAyzPrintable()
Brw BrkAyzPrintable(StrzCurPjf)
End Sub

Private Function BrkAyzRepeat(S) As String()
Dim L$: L = S
Dim T$, J&
While Len(L) > 0
    DoEvents
    T = ShfTermzRepeatedOrNot(L)
'    Debug.Print J, Len(L), Len(T), RepeatSts(T)
'    J = J + 1
    PushI BrkAyzRepeat, T
'    Stop
Wend
End Function
Private Function BrkAyzPrintable(S) As String()
Dim L$: L = S
#If True Then
    While Len(L) > 0
        Push BrkAyzPrintable, ShfTermzPrintable(L)
    Wend
#Else
    Dim T$, J&, I%
    While Len(L) > 0
        DoEvents
        T = ShfTermzPrintable(L)
        S = PrintableSts(T)
        Debug.Print J, Len(L), Len(T), S,
        If S = "NonPrintable" Then
            For I = 1 To Min(Len(T), 10)
                Debug.Print Asc(Mid(T, I, 1)); " ";
            Next
        End If
        Debug.Print
        
        J = J + 1
        PushI BrkAyzPrintable, T
    '    Stop
    Wend
#End If
End Function
Private Function PrintableSts$(T)
Dim IsPrintable As Boolean
IsPrintable = IsAscPrintable(Asc(FstChr(T)))
Dim J&
For J = 2 To Len(T)
    If IsPrintable <> IsAscPrintablezStrI(T, J) Then
        PrintableSts = "Mix"
        Stop
        Exit Function
    End If
Next
PrintableSts = IIf(IsPrintable, "Prt", "Non")
End Function

Private Function RepeatSts$(T)
'If Len(T) = 199 Then Stop
Select Case Len(T)
Case 0: RepeatSts = "ZeroByt": Exit Function
Case 1: RepeatSts = "OneByt":  Exit Function
Case Else
    Dim IsRepeat As Boolean, Las$, C$, IsSam As Boolean
    Las = SndChr(T)
    IsRepeat = FstChr(T) = Las
    Dim J&
    For J = 3 To Len(T)
        C = Mid(T, J, 1)
        IsSam = C = Las
        Select Case True
        Case IsRepeat And IsSam:
        Case IsRepeat: Stop: RepeatSts = "Mixed": Exit Function
        Case IsSam:    Stop: RepeatSts = "Mixed": Exit Function
        Case Else: Las = C
        End Select
    Next
End Select
RepeatSts = IIf(IsRepeat, "Repated", "Dif")
End Function
Private Function ShfTermzRepeatedOrNot$(OStr$)
If OStr = "" Then Exit Function
Dim J&, C$, Las$, IsSam As Boolean, IsRepeat As Boolean
Las = SndChr(OStr)
IsRepeat = FstChr(OStr) = Las
For J = 3 To Len(OStr)
    C = Mid(OStr, J, 1)
    IsSam = C = Las
    Select Case True
    Case IsSam And IsRepeat
    Case IsSam
        ShfTermzRepeatedOrNot = Left(OStr, J - 2)
        OStr = Mid(OStr, J - 1)
        Exit Function
    Case IsRepeat
        ShfTermzRepeatedOrNot = Left(OStr, J - 1)
        OStr = Mid(OStr, J)
        Exit Function
    Case Else
        Las = C
    End Select
Next
ShfTermzRepeatedOrNot = OStr
OStr = ""
End Function

Private Sub BrwRepeatedBytes(S)
Dim J&, B%, B1%, RepeatCnt&, L&
L = Len(S)
If L = 0 Then Exit Sub
B = Asc(FstChr(S)): RepeatCnt = 1
Erase XX
X FmtQQ("Len(?)", L)
For J = 2 To L
    B1 = Asc(Mid(S, J, 1))
    Select Case True
    Case B = B1:        RepeatCnt = RepeatCnt + 1
    Case Else
        If RepeatCnt > 1 Then
            X FmtQQ("Pos(?) Asc(?) RepeatCnt(?)", J, B, RepeatCnt)
            RepeatCnt = 1
        End If
        B = B1
    End Select
Next
Brw SyAddIxPfx(XX)
Erase XX
End Sub

Sub BBB()
If Asc(" ") <> 0 Then Debug.Print 0
If Asc("") <> 1 Then Debug.Print 1
If Asc("") <> 2 Then Debug.Print 2
If Asc("") <> 3 Then Debug.Print 3
If Asc("") <> 4 Then Debug.Print 4
If Asc("") <> 5 Then Debug.Print 5
If Asc("") <> 6 Then Debug.Print 6
If Asc("") <> 7 Then Debug.Print 7
If Asc("") <> 8 Then Debug.Print 8
If Asc("    ") <> 9 Then Debug.Print 9
If Asc("") <> 11 Then Debug.Print 11
If Asc("") <> 12 Then Debug.Print 12
If Asc("") <> 14 Then Debug.Print 14
If Asc("") <> 15 Then Debug.Print 15
If Asc("") <> 16 Then Debug.Print 16
If Asc("") <> 17 Then Debug.Print 17
If Asc("") <> 18 Then Debug.Print 18
If Asc("") <> 19 Then Debug.Print 19
If Asc("") <> 20 Then Debug.Print 20
If Asc("") <> 21 Then Debug.Print 21
If Asc("") <> 22 Then Debug.Print 22
If Asc("") <> 23 Then Debug.Print 23
If Asc("") <> 24 Then Debug.Print 24
If Asc("") <> 25 Then Debug.Print 25
If Asc("") <> 26 Then Debug.Print 26
If Asc("") <> 27 Then Debug.Print 27
If Asc("") <> 28 Then Debug.Print 28
If Asc("") <> 29 Then Debug.Print 29
If Asc("") <> 30 Then Debug.Print 30
If Asc("") <> 31 Then Debug.Print 31
If Asc(" ") <> 32 Then Debug.Print 32
If Asc("!") <> 33 Then Debug.Print 33
If Asc("""") <> 34 Then Debug.Print 34
If Asc("#") <> 35 Then Debug.Print 35
If Asc("$") <> 36 Then Debug.Print 36
If Asc("%") <> 37 Then Debug.Print 37
If Asc("&") <> 38 Then Debug.Print 38
If Asc("'") <> 39 Then Debug.Print 39
If Asc("(") <> 40 Then Debug.Print 40
If Asc(")") <> 41 Then Debug.Print 41
If Asc("*") <> 42 Then Debug.Print 42
If Asc("+") <> 43 Then Debug.Print 43
If Asc(",") <> 44 Then Debug.Print 44
If Asc("-") <> 45 Then Debug.Print 45
If Asc(".") <> 46 Then Debug.Print 46
If Asc("/") <> 47 Then Debug.Print 47
If Asc("0") <> 48 Then Debug.Print 48
If Asc("1") <> 49 Then Debug.Print 49
If Asc("2") <> 50 Then Debug.Print 50
If Asc("3") <> 51 Then Debug.Print 51
If Asc("4") <> 52 Then Debug.Print 52
If Asc("5") <> 53 Then Debug.Print 53
If Asc("6") <> 54 Then Debug.Print 54
If Asc("7") <> 55 Then Debug.Print 55
If Asc("8") <> 56 Then Debug.Print 56
If Asc("9") <> 57 Then Debug.Print 57
If Asc(":") <> 58 Then Debug.Print 58
If Asc(";") <> 59 Then Debug.Print 59
If Asc("<") <> 60 Then Debug.Print 60
If Asc("=") <> 61 Then Debug.Print 61
If Asc(">") <> 62 Then Debug.Print 62
If Asc("?") <> 63 Then Debug.Print 63
If Asc("@") <> 64 Then Debug.Print 64
If Asc("A") <> 65 Then Debug.Print 65
If Asc("B") <> 66 Then Debug.Print 66
If Asc("C") <> 67 Then Debug.Print 67
If Asc("D") <> 68 Then Debug.Print 68
If Asc("E") <> 69 Then Debug.Print 69
If Asc("F") <> 70 Then Debug.Print 70
If Asc("G") <> 71 Then Debug.Print 71
If Asc("H") <> 72 Then Debug.Print 72
If Asc("I") <> 73 Then Debug.Print 73
If Asc("J") <> 74 Then Debug.Print 74
If Asc("K") <> 75 Then Debug.Print 75
If Asc("L") <> 76 Then Debug.Print 76
If Asc("M") <> 77 Then Debug.Print 77
If Asc("N") <> 78 Then Debug.Print 78
If Asc("O") <> 79 Then Debug.Print 79
If Asc("P") <> 80 Then Debug.Print 80
If Asc("Q") <> 81 Then Debug.Print 81
If Asc("R") <> 82 Then Debug.Print 82
If Asc("S") <> 83 Then Debug.Print 83
If Asc("T") <> 84 Then Debug.Print 84
If Asc("U") <> 85 Then Debug.Print 85
If Asc("V") <> 86 Then Debug.Print 86
If Asc("W") <> 87 Then Debug.Print 87
If Asc("X") <> 88 Then Debug.Print 88
If Asc("Y") <> 89 Then Debug.Print 89
If Asc("Z") <> 90 Then Debug.Print 90
If Asc("[") <> 91 Then Debug.Print 91
If Asc("\") <> 92 Then Debug.Print 92
If Asc("]") <> 93 Then Debug.Print 93
If Asc("^") <> 94 Then Debug.Print 94
If Asc("_") <> 95 Then Debug.Print 95
If Asc("`") <> 96 Then Debug.Print 96
If Asc("a") <> 97 Then Debug.Print 97
If Asc("b") <> 98 Then Debug.Print 98
If Asc("c") <> 99 Then Debug.Print 99
If Asc("d") <> 100 Then Debug.Print 100
If Asc("e") <> 101 Then Debug.Print 101
If Asc("f") <> 102 Then Debug.Print 102
If Asc("g") <> 103 Then Debug.Print 103
If Asc("h") <> 104 Then Debug.Print 104
If Asc("i") <> 105 Then Debug.Print 105
If Asc("j") <> 106 Then Debug.Print 106
If Asc("k") <> 107 Then Debug.Print 107
If Asc("l") <> 108 Then Debug.Print 108
If Asc("m") <> 109 Then Debug.Print 109
If Asc("n") <> 110 Then Debug.Print 110
If Asc("o") <> 111 Then Debug.Print 111
If Asc("p") <> 112 Then Debug.Print 112
If Asc("q") <> 113 Then Debug.Print 113
If Asc("r") <> 114 Then Debug.Print 114
If Asc("s") <> 115 Then Debug.Print 115
If Asc("t") <> 116 Then Debug.Print 116
If Asc("u") <> 117 Then Debug.Print 117
If Asc("v") <> 118 Then Debug.Print 118
If Asc("w") <> 119 Then Debug.Print 119
If Asc("x") <> 120 Then Debug.Print 120
If Asc("y") <> 121 Then Debug.Print 121
If Asc("z") <> 122 Then Debug.Print 122
If Asc("{") <> 123 Then Debug.Print 123
If Asc("|") <> 124 Then Debug.Print 124
If Asc("}") <> 125 Then Debug.Print 125
If Asc("~") <> 126 Then Debug.Print 126
If Asc("") <> 127 Then Debug.Print 127
If Asc("Ä") <> 128 Then Debug.Print 128
If Asc("Å") <> 129 Then Debug.Print 129
If Asc("Ç") <> 130 Then Debug.Print 130
If Asc("É") <> 131 Then Debug.Print 131
If Asc("Ñ") <> 132 Then Debug.Print 132
If Asc("Ö") <> 133 Then Debug.Print 133
If Asc("Ü") <> 134 Then Debug.Print 134
If Asc("á") <> 135 Then Debug.Print 135
If Asc("à") <> 136 Then Debug.Print 136
If Asc("â") <> 137 Then Debug.Print 137
If Asc("ä") <> 138 Then Debug.Print 138
If Asc("ã") <> 139 Then Debug.Print 139
If Asc("å") <> 140 Then Debug.Print 140
If Asc("ç") <> 141 Then Debug.Print 141
If Asc("é") <> 142 Then Debug.Print 142
If Asc("è") <> 143 Then Debug.Print 143
If Asc("ê") <> 144 Then Debug.Print 144
If Asc("ë") <> 145 Then Debug.Print 145
If Asc("í") <> 146 Then Debug.Print 146
If Asc("ì") <> 147 Then Debug.Print 147
If Asc("î") <> 148 Then Debug.Print 148
If Asc("ï") <> 149 Then Debug.Print 149
If Asc("ñ") <> 150 Then Debug.Print 150
If Asc("ó") <> 151 Then Debug.Print 151
If Asc("ò") <> 152 Then Debug.Print 152
If Asc("DocOf") <> 153 Then Debug.Print 153
If Asc("ö") <> 154 Then Debug.Print 154
If Asc("õ") <> 155 Then Debug.Print 155
If Asc("ú") <> 156 Then Debug.Print 156
If Asc("ù") <> 157 Then Debug.Print 157
If Asc("û") <> 158 Then Debug.Print 158
If Asc("ü") <> 159 Then Debug.Print 159
If Asc("†") <> 160 Then Debug.Print 160
If Asc("°") <> 161 Then Debug.Print 161
If Asc("¢") <> 162 Then Debug.Print 162
If Asc("£") <> 163 Then Debug.Print 163
If Asc("§") <> 164 Then Debug.Print 164
If Asc("•") <> 165 Then Debug.Print 165
If Asc("¶") <> 166 Then Debug.Print 166
If Asc("ß") <> 167 Then Debug.Print 167
If Asc("®") <> 168 Then Debug.Print 168
If Asc("©") <> 169 Then Debug.Print 169
If Asc("™") <> 170 Then Debug.Print 170
If Asc("´") <> 171 Then Debug.Print 171
If Asc("¨") <> 172 Then Debug.Print 172
If Asc("≠") <> 173 Then Debug.Print 173
If Asc("Æ") <> 174 Then Debug.Print 174
If Asc("Ø") <> 175 Then Debug.Print 175
If Asc("∞") <> 176 Then Debug.Print 176
If Asc("±") <> 177 Then Debug.Print 177
If Asc("≤") <> 178 Then Debug.Print 178
If Asc("≥") <> 179 Then Debug.Print 179
If Asc("¥") <> 180 Then Debug.Print 180
If Asc("µ") <> 181 Then Debug.Print 181
If Asc("∂") <> 182 Then Debug.Print 182
If Asc("∑") <> 183 Then Debug.Print 183
If Asc("∏") <> 184 Then Debug.Print 184
If Asc("π") <> 185 Then Debug.Print 185
If Asc("∫") <> 186 Then Debug.Print 186
If Asc("ª") <> 187 Then Debug.Print 187
If Asc("º") <> 188 Then Debug.Print 188
If Asc("Ω") <> 189 Then Debug.Print 189
If Asc("æ") <> 190 Then Debug.Print 190
If Asc("ø") <> 191 Then Debug.Print 191
If Asc("¿") <> 192 Then Debug.Print 192
If Asc("¡") <> 193 Then Debug.Print 193
If Asc("¬") <> 194 Then Debug.Print 194
If Asc("√") <> 195 Then Debug.Print 195
If Asc("ƒ") <> 196 Then Debug.Print 196
If Asc("≈") <> 197 Then Debug.Print 197
If Asc("∆") <> 198 Then Debug.Print 198
If Asc("«") <> 199 Then Debug.Print 199
If Asc("»") <> 200 Then Debug.Print 200
If Asc("…") <> 201 Then Debug.Print 201
If Asc(" ") <> 202 Then Debug.Print 202
If Asc("À") <> 203 Then Debug.Print 203
If Asc("Ã") <> 204 Then Debug.Print 204
If Asc("Õ") <> 205 Then Debug.Print 205
If Asc("Œ") <> 206 Then Debug.Print 206
If Asc("œ") <> 207 Then Debug.Print 207
If Asc("–") <> 208 Then Debug.Print 208
If Asc("—") <> 209 Then Debug.Print 209
If Asc("“") <> 210 Then Debug.Print 210
If Asc("”") <> 211 Then Debug.Print 211
If Asc("‘") <> 212 Then Debug.Print 212
If Asc("’") <> 213 Then Debug.Print 213
If Asc("÷") <> 214 Then Debug.Print 214
If Asc("◊") <> 215 Then Debug.Print 215
If Asc("ÿ") <> 216 Then Debug.Print 216
If Asc("Ÿ") <> 217 Then Debug.Print 217
If Asc("⁄") <> 218 Then Debug.Print 218
If Asc("€") <> 219 Then Debug.Print 219
If Asc("‹") <> 220 Then Debug.Print 220
If Asc("›") <> 221 Then Debug.Print 221
If Asc("ﬁ") <> 222 Then Debug.Print 222
If Asc("ﬂ") <> 223 Then Debug.Print 223
If Asc("‡") <> 224 Then Debug.Print 224
If Asc("·") <> 225 Then Debug.Print 225
If Asc("‚") <> 226 Then Debug.Print 226
If Asc("„") <> 227 Then Debug.Print 227
If Asc("‰") <> 228 Then Debug.Print 228
If Asc("Â") <> 229 Then Debug.Print 229
If Asc("Ê") <> 230 Then Debug.Print 230
If Asc("Á") <> 231 Then Debug.Print 231
If Asc("Ë") <> 232 Then Debug.Print 232
If Asc("È") <> 233 Then Debug.Print 233
If Asc("Í") <> 234 Then Debug.Print 234
If Asc("Î") <> 235 Then Debug.Print 235
If Asc("Ï") <> 236 Then Debug.Print 236
If Asc("Ì") <> 237 Then Debug.Print 237
If Asc("Ó") <> 238 Then Debug.Print 238
If Asc("Ô") <> 239 Then Debug.Print 239
If Asc("") <> 240 Then Debug.Print 240
If Asc("Ò") <> 241 Then Debug.Print 241
If Asc("Ú") <> 242 Then Debug.Print 242
If Asc("Û") <> 243 Then Debug.Print 243
If Asc("Ù") <> 244 Then Debug.Print 244
If Asc("ı") <> 245 Then Debug.Print 245
If Asc("ˆ") <> 246 Then Debug.Print 246
If Asc("˜") <> 247 Then Debug.Print 247
If Asc("¯") <> 248 Then Debug.Print 248
If Asc("˘") <> 249 Then Debug.Print 249
If Asc("˙") <> 250 Then Debug.Print 250
If Asc("˚") <> 251 Then Debug.Print 251
If Asc("¸") <> 252 Then Debug.Print 252
If Asc("˝") <> 253 Then Debug.Print 253
If Asc("˛") <> 254 Then Debug.Print 254
If Asc("ˇ") <> 255 Then Debug.Print 255
End Sub

