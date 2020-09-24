Attribute VB_Name = "MxVb_Str_Lines_Lbl"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Lines_Lbl."
Private Sub B_LinesLbl()
Dmp LinesLbl(543)
End Sub

Function LinesLblLn$(S)
Const CSub$ = CMod & "LinesLblLn"
If Not IsStr(S) Then Thw CSub, "Given @S is not str", "Tyn(S)", TypeName(S)
LinesLblLn = LinesLbl(WdtLines(S)) & vbCrLf & S
End Function

Function LinesLbl$(Wdt%)
Dim O$()
PushNB O, Ln100(Wdt)
PushNB O, Ln010(Wdt)
PushNB O, Ln001(Wdt)
LinesLbl = JnCrLf(O)
End Function
Private Function Ln001$(Wdt%)
Const C$ = "123456789 "
Dim N&: N = (Wdt \ 10) + 1
Ln001 = Left(StrDup(C, N), Wdt)
End Function
Private Function Ln010$(Wdt%)
If Wdt < 9 Then Exit Function
Dim O$()
    PushI O, Space(9)
    Dim J%: For J = 0 To (Wdt \ 10)
        Dim C$: C = Right(CStr((J Mod 10) + 1), 1)
        PushI O, StrDup(C, 10)
    Next
Ln010 = Left(Jn(O), Wdt)
End Function
Private Function Ln100$(Wdt%)
If Wdt < 99 Then Exit Function
Dim O$()
    PushI O, Space(99)
    Dim J%: For J = 0 To (Wdt \ 100)
        Dim C$: C = Right(CStr((J Mod 10) + 1), 1)
        PushI O, StrDup(C, 99) & " "
    Next
Ln100 = Left(Jn(O), Wdt)
End Function
