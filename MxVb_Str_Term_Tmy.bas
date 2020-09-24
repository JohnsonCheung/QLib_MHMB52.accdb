Attribute VB_Name = "MxVb_Str_Term_Tmy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Term_Tmy."
Function Tmy2(Ln) As String(): Tmy2 = TmyN(Ln, 2): End Function
Function Tmy3(Ln) As String(): Tmy3 = TmyN(Ln, 3): End Function
Function Tmy4(Ln) As String(): Tmy4 = TmyN(Ln, 4): End Function
Function Tmy5(Ln) As String(): Tmy5 = TmyN(Ln, 5): End Function
Function TmyN(Ln, NTm%) As String()
Dim L$: L = Ln
Dim J%: For J = 1 To NTm
    PushI TmyN, ShfTm(L)
Next
End Function

Function DyTmly(Tmly$()) As Variant()
Dim ODy()
Dim Tml: For Each Tml In Itr(Tmly)
    PushI ODy, Tmy(Tml)
Next
DyTmly = DyReDimDc(ODy)
End Function
Function Tmy1r(L) As String(): Tmy1r = Tmynr(L, 1): End Function
Function Tmy2r(L) As String(): Tmy2r = Tmynr(L, 2): End Function
Function Tmy3r(L) As String(): Tmy3r = Tmynr(L, 3): End Function
Function Tmy4r(L) As String(): Tmy4r = Tmynr(L, 4): End Function
Function Tmy5r(L) As String(): Tmy5r = Tmynr(L, 5): End Function
Function Tmynr(L, N%) As String()
Dim S$: S = L
Dim J%: For J = 1 To N
    PushI Tmynr, ShfTm(S)
Next
PushI Tmynr, S
End Function

Function Tm2$(Tml): Tm2 = TmN(Tml, 2): End Function
Function Tm3$(Tml): Tm3 = TmN(Tml, 3): End Function
Function Tm4$(Tml): Tm4 = TmN(Tml, 4): End Function
Private Sub B_TmN()
Dim N%, A$
N = 1: A = "a b c": Ept = "a": GoSub Tst
N = 2: A = "a b c": Ept = "b": GoSub Tst
N = 3: A = "a b c": Ept = "c": GoSub Tst
Exit Sub
Tst:
    Act = TmN(A, N)
    C
    Return
End Sub

Function TmN$(S, N%)
Dim L$, J%
L = LTrim(S)
For J = 1 To N - 1
    L = RmvA1T(L)
Next
TmN = Tm1(L)
End Function

Function AetTml(Tml, Optional C As eCas) As Dictionary: Set AetTml = AetAy(Tmy(Tml), C): End Function
Function TmyTml(Tml) As String():                           TmyTml = Tmy(Tml):           End Function
Function Tmy(Tml) As String() '#Tm-Array#
Dim L$, J%: L = Tml
While L <> ""
    J = J + 1: If J > 500 Then RaiseMsg CSub & ": Looping too much"
    PushI Tmy, ShfTm(L)
Wend
End Function
Function TmlTmy$(Tmy$()): TmlTmy = Tml(Tmy):     End Function
Function NTm%(Tml):          NTm = Si(Tmy(Tml)): End Function

Function Tml$(Tmy$()) '#Tm-Array#
Dim O$()
Dim Tm: For Each Tm In Itr(Tmy)
    PushI O, QuoTm(Tm)
Next
Tml = JnSpc(O)
End Function
Function QuoTm$(Tm)
Select Case True
Case HasSpc(Tm), HasCr(Tm), HasLf(Tm)
    QuoTm = QuoSq(Tm)
Case Else
    QuoTm = Tm
End Select
End Function
Function TmlSrt$(ATml): TmlSrt = Tml(SySrtQ(Tmy(ATml))): End Function
Function TmlAp$(ParamArray ApTm())
Dim Av(): Av = ApTm: TmlAp = TmlAy(Av)
End Function
Function TmyAy(Ay) As String()
Dim V: For Each V In Itr(Ay)
    PushI TmyAy, QuoTm(V)
Next
End Function
Function TmyAp(ParamArray Ap()) As String(): Dim Av(): Av = Ap: TmyAp = TmyAy(Av): End Function
Function TmlAy$(Ay): TmlAy = Tml(TmyAy(Ay)): End Function
Function NmAftTm$(L, Tm$)
Dim S$: S = L
If Not IsShfTm(S, Tm) Then Exit Function
NmAftTm = TakNm(S)
End Function

Private Sub B_ShfTm()
GoSub T1
GoSub T2
GoSub T3
Exit Sub
Dim OLnEpt$, OLn$
T1:
    Ept = "A"
    OLnEpt = "B "
    OLn = "  A   B "
    GoTo Tst
T2:
    OLn = " sdlfkj sdlkf"
    Ept = "sdlfkj"
    OLnEpt = "sdlkf"
    GoSub Tst
T3:
    OLn = "   [ kdjf ] sdlkfj1"
    Ept = " kdjf "
    OLnEpt = "sdlkfj1"
    GoTo Tst
Tst:
    Act = ShfTm(OLn)
    C
    Ass OLnEpt = OLn
    Return
End Sub

Function HasTmo3(L, T1, T2, T3) As Boolean
Dim S$: S = L
Select Case True
Case ShfTm(S) <> T1, ShfTm(S) <> T2, ShfTm(S) <> T3
Case Else: HasTmo3 = True
End Select
End Function

Function HasTmo2(L, T1, T2) As Boolean
Dim S$: S = L
Select Case True
Case ShfTm(S) <> T1, ShfTm(S) <> T2
Case Else: HasTmo2 = True
End Select
End Function
Function TmlMinus$(Tml1$, Tml2$):                  TmlMinus = Tml(SyMinus(Tmy(Tml1), Tmy(Tml2))): End Function
Function TmlMinusTmy$(Tml$, Tmy$()):            TmlMinusTmy = TmlTmy(SyMinus(TmyTml(Tml), Tmy)):  End Function
Function TmyMinusTml(Tmy$(), Tml$) As String(): TmyMinusTml = SyMinus(Tmy, TmyTml(Tml)):          End Function
Function HasTmo1(S, T1) As Boolean:                 HasTmo1 = Tm1(S) = T1:                        End Function
Function HasT2(Ln, T2) As Boolean:                    HasT2 = Tm2(Ln) = T2:                       End Function

Function FnyFF(FF$) As String():       FnyFF = Tmy(FF):               End Function
Function SyMinusSS(Sy$(), Ss$):    SyMinusSS = SyMinus(Sy, SySs(Ss)): End Function
Function FnyMinusFF(Fny$(), FF$): FnyMinusFF = TmyMinusTml(Fny, FF):  End Function
Function Tm2y(Tmly) As String()
Dim Tml: For Each Tml In Itr(Tmly)
    PushI Tm2y, Tm2(Tml)
Next
End Function
Function Tm3y(Tmly) As String()
Dim Tml: For Each Tml In Itr(Tmly)
    Stop 'PushI Tm3y, Tm3(L)
Next
End Function

Function NRstTmy(L, N%) As Variant() '#NTm-and-Rest# with @N+1 ele from @L
Dim S$: S = L
Dim J%: For J = 1 To N
    PushI NRstTmy, ShfTm(S)
Next
PushI NRstTmy, S
End Function

Private Sub TmyyTmll__Tst()
GoSub T1
Exit Sub
Dim Tmll$, Act()
Const Tmll1$ = "[WMB52Pthi $] [WMB52PthCpy2 $] [WMB52PthCpy1 $] [WMB52IsCpy1 [] [ As Boolean]] [MB52IsCPy2 [] [ As Boolean]]" & _
" [WPhFxi $]" & _
" [WSkuFxi $] [WSkuPthCpyTo $] [WSkuIsCpyTo [] [ As Boolean]]" & _
" [WZHT0Pthi $] [ZHT0Fxi $] [WZHT0Fxw $]"
T1: Tmll = Tmll1: GoTo Tst
Tst: Act = TmyyTmll(Tmll1): BrwLndy Act: Stop: Return
End Sub
Function TmyyTmll(Tmll$) As Variant()
Dim Tmly$(): Tmly = Tmy(Tmll)
Dim Tml: For Each Tml In Itr(Tmly)
    PushI TmyyTmll, Tmy(Tml)
Next
End Function
