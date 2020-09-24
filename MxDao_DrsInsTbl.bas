Attribute VB_Name = "MxDao_DrsInsTbl"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_DrsInsTbl."

Sub InsTblDrs(D As Database, T, B As Drs)
Dim F$(): F = AyIntersect(Fny(D, T), B.Fny)
InsRsDy RsTFny(D, T, F), DrsSelFny(B, F).Dy
End Sub

Sub InsTblDy(D As Database, T, Dy())
InsRsDy RsTbl(D, T), Dy
End Sub
