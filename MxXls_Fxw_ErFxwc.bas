Attribute VB_Name = "MxXls_Fxw_ErFxwc"
Option Compare Text
Option Explicit

Function WarnFxwcShdNB(Fx$, W$, Coln$) As String()
If HasFxwcBlnk(Fx, W, Coln) Then
    PushI WarnFxwcShdNB, FmtQQ("There are blank value in column[?] of Ws[?] of Excel[?], these rows will be ignored", Coln, W, Fx)
End If
End Function
Function EryFxwcShdAllInAp(Fx$, W$, Coln$, ParamArray ApIn()) As String()
Dim Plnt$()
    Plnt = DcStrDisFxw(Fx, W, "Plant")

End Function
