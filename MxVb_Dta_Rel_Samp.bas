Attribute VB_Name = "MxVb_Dta_Rel_Samp"
Option Compare Text
Const CMod$ = "MxVb_Dta_Rel_Samp."
Option Explicit

Function sampRel() As Dictionary
Set sampRel = Rel(sampLyRel)
End Function
Function sampLyRel() As String()
Erase XX
X "A B"
X "B A"
sampLyRel = XX
Erase XX
End Function
