Attribute VB_Name = "MxDao_Sql_QpGp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_QpGp."

Function QpGp$(Gp$):                      QpGp = StrPfxIfNB(Gp, C_Gp):               End Function
Function QpGpVblyEpr$(VblyEpr$()): QpGpVblyEpr = C_Gp & JnCrLf(FmtVblyEpr(VblyEpr)): End Function

Function QpGpFf$(FF$): QpGpFf = vbCrLf & " Group By " & QpFf(FF): End Function

Private Sub B_QpGpVblyEpr()
Dim VblyEpr$()
    Push VblyEpr, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push VblyEpr, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push VblyEpr, "3sldkfjsdf"
DmpAy SplitVBar(QpGpVblyEpr(VblyEpr))
End Sub
