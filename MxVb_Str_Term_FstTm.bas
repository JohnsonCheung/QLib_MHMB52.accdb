Attribute VB_Name = "MxVb_Str_Term_FstTm"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Term_FstTm."
Function BrkTm1(L$) As S12
Dim P%: P = WPosT1LasChr(L)
With BrkTm1
    .S1 = RmvBktSq(Left(L, P))
    .S1 = LTrim(Mid(L, P))
End With
End Function
Function IsShfTm(OLn$, Tm$) As Boolean: IsShfTm = ShfTm(OLn) = Tm: End Function
Function TakTm$(L)
Dim P%: P = WPosT1LasChr(L)
Dim O$: O = Left(L, P)
TakTm = RmvBktSq(O)
End Function
Function ShfTm$(OLn$)
Dim P%: P = WPosT1LasChr(OLn)
Dim O$: O = Left(OLn, P)
ShfTm = RmvBktSq(O)
OLn = LTrim(Mid(OLn, P + 1))
End Function
Private Function WPosT1LasChr%(L)
Dim O%
If ChrFst(L) = "[" Then
    O = PosBktCls(L, 1, "["): If O = 0 Then Thw CSub, "With [, but no ]", "Ln", L
Else
    O = InStr(L, " ")
    If O = 0 Then
        O = Len(L)
    Else
        O = O - 1
    End If
End If
WPosT1LasChr = O
End Function

Private Sub B_Tml()
GoSub T1
GoSub T2
GoSub T3
Exit Sub
Dim Ln$, LnEpt$
T1: Ln = "  [ksldfj ]    a": Ept = "ksldfj ":  LnEpt = "a": GoSub Tst
T2: Ln = "  [ ksldfj ]   a": Ept = " ksldfj ": LnEpt = "a": GoSub Tst
T3: Ln = "  [ksldfj]    a":  Ept = "ksldfj":   LnEpt = "a": GoSub Tst
Exit Sub
Tst:
    Act = ShfTm(Ln)
    C
    If Ln <> LnEpt Then Stop
    Return
End Sub
Function RmvTm1$(L)
Dim OLn$: OLn = L
ShfTm OLn
RmvTm1 = OLn
End Function
Function Tm1y(Tmly$()) As String(): Tm1y = Tm1yAy(Tmly): End Function
Function Tm1yAy(Ay) As String()
Dim Tml: For Each Tml In Itr(Ay)
    PushI Tm1yAy, Tm1(Tml)
Next
End Function
Function Tm1$(Tml): Tm1 = ShfTm(CStr(Tml)): End Function
Function RplT1$(L, T1$, By$)
If HasTmo1(L, T1) Then
    RplT1 = By & " " & RmvA1T(L)
Else
    RplT1 = L
End If
End Function
