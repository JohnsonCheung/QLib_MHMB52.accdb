Attribute VB_Name = "MxVb_Bfr"
Option Compare Text
Option Explicit
Private X$()
Const CMod$ = "MxVb_Bfr."
Private Bfr$()
Sub ClrBfr():                           Erase Bfr:                                              End Sub
Sub BfrLn():                            PushI Bfr, "":                                          End Sub
Sub BfrBox(S$, Optional C$ = "*"):      PushIAy Bfr, Box(S, C):                                 End Sub
Sub BfrULin(S$, Optional ChrUL$ = "-"): PushI Bfr, S: PushI Bfr, StrDup(ChrFst(ChrUL), Len(S)): End Sub
Sub BrwBfr():                           BrwAy Bfr:                                              End Sub
Function LyBfr() As String(): LyBfr = Bfr: Erase Bfr: End Function
Function LinesBfr$(): LinesBfr = JnCrLf(LyBfr): End Function
Sub BfrV(Optional V)
If IsEmpty(V) Then PushI Bfr, "": Exit Sub
PushIAy Bfr, FmtVal(V)
End Sub
Sub BfrTab(V)
If IsArray(V) Then
    BfrV AmTab(V)
Else
    BfrV vbTab & V
End If
End Sub
