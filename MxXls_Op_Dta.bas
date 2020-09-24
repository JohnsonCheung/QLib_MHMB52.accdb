Attribute VB_Name = "MxXls_Op_Dta"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Op_Dta."
Function RgAyV(Ay, At As Range) As Range: Set RgAyV = RgSq(SqCol(Ay), At): End Function
Function RgAyH(Ay, At As Range) As Range: Set RgAyH = RgSq(SqRow(Ay), At): End Function
