Attribute VB_Name = "MxVb_Str_QuoFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Quo."
Function BrkQuo(QuoStr$) As S12
Dim L%: L = Len(QuoStr)
Dim S1$, S2$
Select Case L
Case 0:
Case 1
    S1 = QuoStr
    S2 = QuoStr
Case 2
    S1 = Left(QuoStr, 1)
    S2 = Right(QuoStr, 1)
Case Else
    If InStr(QuoStr, "*") > 0 Then
        BrkQuo = Brk(QuoStr, "*", NoTrim:=True)
        Exit Function
    End If
    Stop
End Select
BrkQuo = S12(S1, S2)
End Function

Function AyQuo(Ay, QuoStr$) As String()
Dim P$, S$
With BrkQuo(QuoStr)
    P = .S1
    S = .S2
End With
AyQuo = AmAddPfxSfx(Ay, P, S)
End Function

Function VstrUnquo$(VstrQuod): VstrUnquo = Replace(RmvFstLas(VstrQuod), vbQuoDbl2, vbQuoDbl): End Function
Function VstrQuo$(Vstr)
':VstrQuo: #Quoted-Vb-Str# ! a str with fst and lst chr is vbQuoDbl and inside each vbQuoDbl is in pair, which will cv to one vbQuoDbl  @@
VstrQuo = vbQuoDbl & Replace(Vstr, vbQuoDbl, vbQuoDbl2) & vbQuoDbl
End Function

Function Quo$(S, QuoStr$)
With BrkQuo(QuoStr)
    Quo = .S1 & S & .S2
End With
End Function

Function QuoIfNB$(IfNB, QuoStr$)
If Trim(IfNB) = "" Then Exit Function
With BrkQuo(QuoStr)
    QuoIfNB = .S1 & IfNB & .S2
End With
End Function

Function QuoVstr$(S):      QuoVstr = vbQuoDbl & Replace(S, vbQuoDbl, vbQuoDbl2) & vbQuoDbl: End Function 'Quote @S as vbStr, which is quoting with double-quote and inside-double-quote will become 2 double-quote.
Function QuoBigBkt$(S):  QuoBigBkt = "{" & S & "}":                                         End Function
Function QuoBkt$(S):        QuoBkt = "(" & S & ")":                                         End Function
Function QuoDte$(S):        QuoDte = QuoBy(S, "#"):                                         End Function
Function QuoDot$(S):        QuoDot = QuoBy(S, "."):                                         End Function
Function QuoBig$(S):        QuoBig = vbBktOpnBig & S & vbBktClsBig:                         End Function
Function QuoDbl$(S):        QuoDbl = QuoBy(S, vbQuoDbl):                                    End Function
Function QuoSng$(S):        QuoSng = QuoBy(S, "'"):                                         End Function
Function QuoSpc$(S):        QuoSpc = QuoBy(S, " "):                                         End Function
Function QuoSq$(S):          QuoSq = "[" & S & "]":                                         End Function
Function QuoBy$(S, By$):     QuoBy = By & S & By:                                           End Function
Function QuoSqIf$(S)
If Not IsNm(S) Then
    QuoSqIf = QuoSq(S)
Else
    QuoSqIf = S
End If
End Function
Function QuoSqAv(Av()) As String() 'Quote each element in @Av by square bracket and return it as string array.
Dim I: For Each I In Av
    PushI QuoSqAv, QuoSq(I)
Next
End Function
