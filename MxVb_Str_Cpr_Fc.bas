Attribute VB_Name = "MxVb_Str_Cpr_Fc"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Cpr_Fc."
Private Sub B_FcLines()
FcLines "A", "A"
End Sub
Sub FcLines(LinesA$, LinesB$)
If LinesA = LinesB Then Exit Sub
BrwStr LinesFc(LinesA, LinesB)
End Sub
Function LinesFcLy$(LyA$(), LyB$()): LinesFcLy = LinesFc(JnCrLf(LyA), JnCrLf(LyB)): End Function
Function LinesFc$(LinesA$, LinesB$):
WrtAB:
    Dim P$: P = PthTmpInst("FileCompare")
    WrtStr LinesA, P & "a.txt"
    WrtStr LinesB, P & "b.txt"
ShellScript:
    Dim Cdl$: Cdl = LinesApLn( _
    FmtQQ("Cd ""?""", P), _
    "Fc a.txt b.txt /N >Fc.txt")
    Shell FcmdCrt("FileCompare", Cdl), vbHide
ChkWaitFfn P & "Fc.txt"
LinesFc = LinesApLn( _
    "In Pth[" & P & "]", _
    "File a.txt lines count [" & NLn(LinesA) & "]", _
    "File b.txt lines count [" & NLn(LinesB) & "]", _
    LinesEndTrim(LinesFt(P & "Fc.txt")))
End Function
