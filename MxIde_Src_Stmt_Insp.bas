Attribute VB_Name = "MxIde_Src_Stmt_Insp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Stmt_Insp."

Sub Insp(Fun$, Msg$, ParamArray Nap())
Dim Nav(): Nav = Nap
BrwAy MsgyFMNav(Fun, "Inspect: " & Msg, Nav)
End Sub
