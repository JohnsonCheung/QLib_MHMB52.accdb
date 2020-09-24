Attribute VB_Name = "MxHtml_QuoHtml"
Option Compare Text
Option Explicit

Function QHtml$(HInr):                           QHtml = QHTagLf(HInr, "html"):                                      End Function
Function QHTagLf$(HInr, Tag$, Optional Atr$):  QHTagLf = WTagB(Tag, Atr) & vbCrLf & HInr & vbCrLf & WTagE$(Tag):     End Function
Function QHTag$(HInr, Tag$, Optional Atr$):      QHTag = WTagB(Tag, Atr) & HInr & WTagE(Tag):                        End Function
Function QHTable$(HInr, Optional Atr$):        QHTable = QHTagLf(HInr, "table", Atr):                                End Function
Function QHTr$(HInr):                             QHTr = QHTag(HInr, "tr"):                                          End Function
Function QHTd$(HInr):                             QHTd = QHTag(HInr, "td"):                                          End Function
Function QHInpTxt(TxtVal$):                   QHInpTxt = QHTag("", "input", FmtQQ("type='text' value='?'", TxtVal)): End Function
Function HTTrAp$(ParamArray Ap())
Dim Av(): Av = Ap: HTTrAp = HTTrDr(Av)
End Function
Function HTTrDr$(Dr)
Dim O$()
Dim V: For Each V In Itr(Dr)
    PushI O, QHTd(V)
Next
HTTrDr = QHTr(Jn(O))
End Function
Private Function WTagB$(Tag$, Optional Atr$)
If Atr = "" Then
    WTagB = "<" & Tag & ">"
Else
    WTagB = "<" & Tag & " " & Atr & ">"
End If
End Function
Private Function WTagE$(Tag$): WTagE = "</" & Tag & ">": End Function
Function HTTrBrkSpc$(L)
With BrkSpc(L)
    HTTrBrkSpc = HTTrAp(QHTd(.S1), QHTd(.S2))
End With
End Function
Function QHCd$(HInr):                              QHCd = QHTagLf(HInr, "code"):                      End Function
Function QHDiv$(HInr, Optional Atr$):             QHDiv = QHTag(HInr, "div", Atr):                    End Function
Function QHScript$(HInr, Optional Atr$):       QHScript = QHTagLf(HInr, "script", Atr):               End Function
Function HTScriptFjs$(Fjs$):                HTScriptFjs = QHScript(LinesFt(Fjs)):                     End Function
Function HTScriptFjsSrc$(Fjs$):          HTScriptFjsSrc = QHTag("", "script", FmtQQ("Src='?'", Fjs)): End Function
