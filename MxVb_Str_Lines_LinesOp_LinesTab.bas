Attribute VB_Name = "MxVb_Str_Lines_LinesOp_LinesTab"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Cml_Tab."

Function LinesTab4Spc$(L$):            LinesTab4Spc = JnCrLf(LyTab4Spc(SplitCrLf(L))): End Function
Function LyTab4Spc(Ly$()) As String():    LyTab4Spc = AmAddPfx(Ly, vbSpc4):            End Function
Function LinesTab$(L$):                    LinesTab = JnCrLf(LyTab(SplitCrLf(L))):     End Function
Function LyTab(Sy$()) As String():            LyTab = AmAddPfx(Sy, vbTab):             End Function
