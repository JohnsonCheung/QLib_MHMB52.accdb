Attribute VB_Name = "MxVb_Str_Lines_Inf"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Lines_Inf."
Function LinesInf$(Lines$)
LinesInf = FmtQQ("Cnt-Si(?-?)", NLn(Lines), Len(Lines))
End Function

Function LinesInfLy$(Ly$())
LinesInfLy = LinesInf(JnCrLf(Ly))
End Function
