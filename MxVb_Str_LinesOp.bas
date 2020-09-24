Attribute VB_Name = "MxVb_Str_LinesOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Lsy."

Function IsLinesMoreThan1Ln(Lines) As Boolean: IsLinesMoreThan1Ln = HasSsub(Lines, vbCrLf): End Function

Function LsyBeiy(Ly$(), B() As Bei) As String()
Dim J&: For J = 0 To UbBei(B)
    PushI LsyBeiy, LinesAwBei(Ly, B(J))
Next
End Function
Function LinesLyNB$(Ly$()): LinesLyNB = JnCrLf(AwNB(Ly)): End Function
Function LinesApLn$(ParamArray ApLn())
Dim Av(): Av = ApLn
LinesApLn = Jn(Av, vbCrLf)
End Function
Function LinesApLnNB$(ParamArray ApLines()) ' Empty Lines will be skipped
Dim Av: Av = ApLines
Av = AwNB(Av)
LinesApLnNB = JnCrLf(Av)
End Function
Function LinesNmLines$(Nm$, Lines$): LinesNmLines = LinesUL(Nm) & vbCrLf & Lines: End Function
Function LinesMix$(A$, B$)
If A = B Then
    LinesMix = A
Else
    LinesMix = LinesApLn(LinesUL(A), B)
End If
End Function

Function LinesStrPfxTab$(Lines): LinesStrPfxTab = JnCrLf(AmAddPfxTab(SplitCrLf(Lines))):          End Function
Function LstsLines$(Lines$):          LstsLines = FmtQQ("-NLn ? -Len ?", NLn(Lines), Len(Lines)): End Function
Function LinesAdd$(A$, B$)
If LTrim(A) = "" Then LinesAdd = B: Exit Function
If LTrim(B) = "" Then LinesAdd = A: Exit Function
LinesAdd = A & vbCrLf & B
End Function
Function WdtLines%(Lines): WdtLines = AyWdt(SplitCrLf(Lines)): End Function

Private Sub B_LinesLasN()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
Debug.Print LinesLasN(A, 3)
End Sub
Function LinesLasN$(Lines$, N%): LinesLasN = JnCrLf(AwLasN(SplitCrLf(Lines), N)): End Function
Function LnFst$(Lines$):             LnFst = BefOrAll(Lines, vbCrLf):             End Function
Function LnLas$(Lines$):             LnLas = EleLas(SplitCrLf(Lines)):            End Function

Private Sub B_LinesEndTrim()
GoSub T1
Exit Sub
Dim Lines$
T1:
    Lines = LinesVbl("lksdf|lsdfj|||")
    GoTo Tst
T2:
    Lines = WWSampLines1
    GoTo Tst
Tst:
    Act = LinesEndTrim(Lines)
    Debug.Print Act & "<"
    Return
Stop
End Sub
Function EixNB&(Ly$())
Dim Eix&: For Eix = UB(Ly) To 0 Step -1
    If Trim(Ly(Eix)) <> "" Then EixNB = Eix: Exit Function
Next
EixNB = -1
End Function
Function LyEndTrim(Ly$()) As String():    LyEndTrim = AyReDim(Ly, EixNB(Ly)):              End Function
Function LinesEndTrim$(Lines):         LinesEndTrim = JnCrLf(LyEndTrim(SplitCrLf(Lines))): End Function
Function NLnMax&(Lsy$())
Dim O&, Lines: For Each Lines In Itr(Lsy)
    O = Max(O, NLn(Lines))
Next
NLnMax = O
End Function

Function LsyLyy(Lyy()) As String()
Dim Ly: For Each Ly In Itr(Lyy)
    PushI LsyLyy, JnCrLf(Ly)
Next
End Function

Private Sub B_NLn()
GoSub ZZ
Exit Sub
ZZ:
    MsgBox NLn(SrclPC) & " " & Si(SplitCrLf(SrclPC))
    Return
End Sub

Function NLn&(Lines)
#If True Then
    NLn = Si(SplitCrLf(Lines))
#Else
    NLn = NSsubzRx(Lines, WWRxNLn)
    If NLn = 0 Then If Lines <> "" Then NLn = 1
#End If
End Function
Private Function WWRxNLn() As RegExp
Static R As RegExp: If IsNothing(R) Then Set R = Rx("/\n/gi")
Set WWRxNLn = R
End Function

Private Function WWSampLsy() As String()
ClrBfr
BfrV WWSampLines1
BfrV WWSampLines2
BfrV WWSampLines3
WWSampLsy = LyBfr
End Function
Private Function WWSampLines1$(): WWSampLines1 = RplVbl("sdklf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf"):                    End Function
Private Function WWSampLines2$(): WWSampLines2 = RplVbl("sdklf2-49230  sdfjldf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf"):    End Function
Private Function WWSampLines3$(): WWSampLines3 = RplVbl("sdsdfklf2-49230  sdfjldf|lskdjflsdf|lsdkjflsdkfjsdflsdf|skldfjdsf|dklfsjdlksjfsldkf"): End Function
