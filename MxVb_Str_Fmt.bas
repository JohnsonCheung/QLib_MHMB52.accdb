Attribute VB_Name = "MxVb_Str_Fmt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Fmt."

Private Sub B_FmtQQAv(): Debug.Print FmtQQ("klsdf?sdf?dsklf", 2, 1): End Sub
Function FmtStr$(BigBktIxVbl$, ParamArray Ap())
'@BigBktIxVbl :Vbl with Ssub {0}...
Dim S$: S = Replace(BigBktIxVbl, "|", vbCrLf)
Dim I, J%: For Each I In Ap
    S = Replace(S, "{" & J & "}", Nz(I, "Null"))
    J = J + 1
Next
FmtStr = S
End Function
Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function
Function FmtQQAv$(QQVbl$, Av())
Dim O$: O = Replace(QQVbl, "|", vbCrLf)
Dim P&: P = 1
Dim I: For Each I In Av
    P = InStr(P, O, "?")
    If P = 0 Then Exit For
    O = Left(O, P - 1) & Replace(O, "?", I, Start:=P, Count:=1)
    P = P + Len(I)
Next
FmtQQAv = O
End Function

Function LinesAy$(Ay): LinesAy = JnCrLf(LyAy(Ay)): End Function
Function LyAy(Ay) As String()
Dim V: For Each V In Itr(Ay)
    PushI LyAy, StrV(V)
Next
End Function

Function StrV$(V)
On Error GoTo X
StrV = V
Exit Function
X: StrV = "Tyn(" & TypeName(V) & ")"
End Function

Function LyV(V, Optional IsAddIx As Boolean, Optional W% = 100, Optional Zer As eZer) As String()
LyV = SplitCrLf(LinesV(V, IsAddIx, W, Zer))
End Function
Function LinesV$(V, Optional IsAddIx As Boolean, Optional W% = 100, Optional Zer As eZer)
Dim O$
Select Case True
Case IsAet(V): Stop ' O = WrdlLines(TmlAy(DikyStr(CvAet(V))), W)
Case IsPrim(V): O = V
Case IsSy(V)
    If IsAddIx Then
        O = JnCrLf(AmAddIxPfx(CvSy(V)))
    Else
        O = JnCrLf(V)
    End If
Case IsNothing(V): O = "#Nothing#"
Case IsEmpty(V):   O = "#Empty"
Case IsMissing(V): O = "#Missing"
Case IsObject(V):  O = "#Obj(" & TypeName(V) & ")"
Case IsNumeric(V): O = FmtNum(V, Zer)
Case IsBool(V):    O = IIf(V, "True", "")
Case IsDate(V):    O = V
Case IsStr(V):     O = LeftOrAll(V, W)
Case IsPrimy(V):   O = FmtPrimy(V, W)
Case IsDi(V):      O = JnCrLf(FmtDi(CvDi(V)))
Case IsObject(V):  O = "#Obj:" & TypeName(V)
Case IsErObj(V):   O = "#Er#"
Case IsEmp(V), IsNull(V)
Case IsArray(V):  O = FmtAy(V, W, IsAddIx)
Case Else:        O = V
End Select
LinesV = O
End Function
Private Function FmtAy$(Ay, W%, IsAddIx As Boolean)
If Si(Ay) = 0 Then FmtAy = "(*Si=0)": Exit Function
If IsAddIx Then
    FmtAy = JnCrLf(AmAddIxPfx(Ay))
Else
    Dim O$()
    Dim I: For Each I In Itr(Ay)
        PushI O, LeftOrAll(I, W)
    Next
    FmtAy = JnCrLf(O)
End If
End Function
Function FmtNum$(Num, Zer As eZer)
If Zer = eZerHid Then
    If Num = 0 Then Exit Function
End If
FmtNum = Num
End Function
Private Function FmtPrimy$(Primy, W%): FmtPrimy = LeftOrAll("*[" & Si(Primy) & "] " & JnSpc(Primy), W): End Function
