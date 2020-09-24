Attribute VB_Name = "MxVb_Str_Lines_BldLines"
Option Compare Text
Option Explicit

Function LinesQVbarNN$(QVbar$, NN$)
Dim O$()
Dim N: For Each N In ItrSS(NN)
    PushI O, RplQVbar(QVbar, N)
Next
LinesQVbarNN = JnCrLf(O)
End Function

Function LinesMacroTmll$(Macro$, Tmll$): LinesMacroTmll = LinesMacroDy(Macro, TmyyTmll(Tmll)): End Function
Function LinesMacroDy$(Macro$, Dy())
Dim O$()
Dim Dr: For Each Dr In Itr(Dy)
    PushI O, FmtMacroAy(Macro, Dr)
Next
LinesMacroDy = JnCrLf(O)
End Function
Function RplMacro(Macro$, Sy$())

End Function
