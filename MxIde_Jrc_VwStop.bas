Attribute VB_Name = "MxIde_Jrc_VwStop"
Option Compare Text
Option Explicit
Private Sub B_VwStop():            VwStop:                         End Sub
Sub VwStop(Optional PatnssAndMd$): VwHtml WFhtmlStop(PatnssAndMd): End Sub
Private Function WFhtmlStop$(PatnssAndMd$)
Dim O$: O = PthTmpFdr("HtmlStop") & "VwStop.html"
WrtStr QHtml(WtmlJrc(JrcyPatnPC("Stop", PatnssAndMd, eULNo))), O, OvrWrt:=True
WFhtmlStop = O
End Function
Private Function WtmlJrc$(JrcyPatnPC$()): WtmlJrc = QHtml(WTable(JrcyPatnPC) & vbCrLf & WScript):                 End Function
Private Function WTable$(JrcyPatnPC$()):   WTable = QHCd(QHTable(JnCrLf(WTry(JrcyPatnPC)), "style='color:red'")): End Function
Private Function WTry(JrcyPatnPC$()) As String()
Dim J&, U&: U = UB(JrcyPatnPC)
Dim L: For Each L In Itr(JrcyPatnPC)
    If J > 10 Then Exit Function
    If J = U Then Exit Function
    J = J + 1
    PushI WTry, WTr(J, L)
Next
End Function
Private Function WTr$(Ix&, Jrcy)
With BrkSpc(Jrcy)
WTr = QHTr(Jn(Array(QHTd(Ix), WTdJmp(.S1), WTdSrcln(.S2))))
End With
End Function
Private Function WTdJmp$(S1$):     WTdJmp = QHTd(QHInpTxt(S1)): End Function
Private Function WTdSrcln$(S2$): WTdSrcln = QHTd(RmvFst(S2)):   End Function
'Private Function WScript$(): WScript = HTScriptFjs("C:\users\user\desktop\a.js"): End Function
Private Function WScript$(): WScript = HTScriptFjsSrc("a.js"): End Function
