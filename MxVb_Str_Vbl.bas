Attribute VB_Name = "MxVb_Str_Vbl"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Vbl."

Function VblLines$(Lines$)
VblLines = Replace(RmvCr(Lines), vbLf, "|")
End Function

Function DicVbl(Vbl$, Optional JnSep$ = vbCrLf) As Dictionary
Set DicVbl = DiLines(SplitVBar(Vbl), JnSep)
End Function

Function SyVbl(Vbl) As String(): SyVbl = SplitVBar(Vbl): End Function

Function LinesVbl$(Vbl)
LinesVbl = Replace(Vbl, "|", vbCrLf)
End Function

Function IsVbl(V) As Boolean
Select Case True
Case Not IsStr(V)
Case HasSsub(V, vbCr)
Case HasSsub(V, vbLf)
Case Else: IsVbl = True
End Select
End Function

Function IsVbly(V) As Boolean
Dim Vbl: For Each Vbl In Itr(V)
    If Not IsVbl(Vbl) Then Exit Function
Next
IsVbly = True
End Function


Function DyVbly(A$()) As Variant()
Dim I
For Each I In Itr(A)
    PushI DyVbly, AmTrim(SplitVBar(I))
Next
End Function
Function DySSVBL(SSVbl$) As Variant()
Dim Ss: For Each Ss In Itr(SplitVBar(SSVbl))
    PushI DySSVBL, SySs(Ss)
Next
End Function

Private Sub B_DyVbly()
Dim VblLy$()
GoSub T1
Exit Sub
T0:
    ClrBfr
    BfrV "1 | 2 | 3"
    BfrV "4 | 5 6 |"
    BfrV "| 7 | 8 | 9 | 10 | 11 |"
    BfrV "12"
    VblLy = LyBfr
    Ept = Array(SySs("1 2 3"), Sy("4", "5 6", ""), Sy("", "7", "8", "9", "10", "11", ""), Sy("12"))
    GoTo Tst
Exit Sub
T1:
    ClrBfr
    BfrV "|lskdf|sdlf|lsdkf"
    BfrV "|lsdf|"
    BfrV "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
    BfrV "|sdf"
    VblLy = LyBfr
    Ept = ""
    GoTo Tst
Tst:
    Act = DyVbly(VblLy)
    DmpDy CvAv(Act)
'    C
    Return
End Sub

Function LyVbl(Vbl) As String()
LyVbl = SplitVBar(Vbl)
End Function
