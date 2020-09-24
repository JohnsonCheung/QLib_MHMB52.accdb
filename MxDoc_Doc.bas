Attribute VB_Name = "MxDoc_Doc"
Option Compare Text
Option Explicit
Const CMod$ = "MxDoc_Doc."
Sub EdtDoc():                                 EdtDocP CPj:            End Sub
Sub EdtDocP(P As VBProject):                  VcFt FfnDocP(P):        End Sub
Function FfnDocP$(P As VBProject):  FfnDocP = PthAssP(P) & "Doc.txt": End Function
Function FfnDocPC$():              FfnDocPC = FfnDocP(CPj):           End Function
