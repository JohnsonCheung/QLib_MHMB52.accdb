Attribute VB_Name = "MxVb_Dta_Rx"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Ssub_Rx."

Function NMch&(S, Patn$):                        NMch = NMchRx(S, Rx(Patn)): End Function
Function NMchRx&(S, R As RegExp):              NMchRx = Mchcoll(S, R).Count: End Function
Function P12Patn(Ln, Patn$) As P12:           P12Patn = P12Rx(Ln, Rx(Patn)): End Function
Function IsMch(S, Rx As RegExp) As Boolean:     IsMch = Rx.Test(S):          End Function
Function IsMchPatn(S, Patn$) As Boolean:    IsMchPatn = IsMch(S, Rx(Patn)):  End Function
Function RxFlg(R As RegExp, FlgMIG$) ' return a new @Rx by modifying the the flags
Dim O As New RegExp
O.Pattern = R.Pattern
SetRxMIG O, FlgMIG
Set RxFlg = O
End Function
Sub SetRxMIG(R As RegExp, FlgMIG$)
With R
    .MultiLine = HasSsub(FlgMIG, "M")
   .IgnoreCase = HasSsub(FlgMIG, "I")
   .Global = HasSsub(FlgMIG, "G")
End With
End Sub
Function CvRx(A) As RegExp: Set CvRx = A: End Function
Function Rxay(TmlPatnFlg$) As RegExp()
Dim PatnFlg: For Each PatnFlg In Itr(Tmy(TmlPatnFlg))
    PushObj Rxay, Rx(CStr(PatnFlg))
Next
End Function
Function RxClone(R As RegExp) As RegExp
Dim O As New RegExp
With O
    .Global = R.Global
    .IgnoreCase = R.IgnoreCase
    .MultiLine = R.MultiLine
    .Pattern = R.Pattern
End With
Set RxClone = O
End Function
Function RxGlobal(R As RegExp) As RegExp
Set RxGlobal = RxClone(R)
RxGlobal.Global = True
End Function
Function RxAyPatnss(Patnss$) As RegExp()
Dim Patn: For Each Patn In Itr(Tmy(Patnss))
    PushObj RxAyPatnss, Rx(Patn)
Next
End Function
Function Rx(PatnFlg) As RegExp
Const CSub$ = CMod & "Rx"
Dim FlgMIG$
Dim Patn$
    If ChrFst(PatnFlg) = "/" Then
        Dim P%: P = InStrRev(PatnFlg, "/"): If P <= 2 Then Thw CSub, "Invalid @PatnFlg: Fst Chr is / but not second / found", "@PatnFlg", PatnFlg
        Dim N%: N = P - 2
        FlgMIG = Mid(PatnFlg, P + 1)
        Patn = Mid(PatnFlg, 2, N)
    Else
        Patn = PatnFlg
    End If
Dim O As New RegExp
O.Pattern = Patn
SetRxMIG O, FlgMIG
Set Rx = O
End Function

Private Sub B_RplPatn()
GoSub T1
Exit Sub
Dim S, By$, Patn$
T1:
    S = "a men is male"
    By = "$1male$3"
    Patn = "(.+)(m[ae]n)(.+)"
    Ept = "a male is male"
    GoTo Tst
Tst:
    Act = RplPatn(S, Patn, By)
    C
    Return
End Sub
Function RplRx$(S, R As RegExp, By$):   RplRx = R.Replace(S, By):           End Function
Function RplPatn$(S, Patn$, ByPatn$): RplPatn = RplRx(S, Rx(Patn), ByPatn): End Function

Private Sub B_P12Rx()
GoSub T1
Exit Sub
Dim Act As P12, Ept As P12, R As RegExp, S
T1:
    S = "aAAa"
    Set R = Rx("Aa")
    Ept = P12(2, 3)
    GoTo Tst
Tst:
    Act = P12Rx(S, R)
    Ass IsEqP12(Act, Ept)
    Return
End Sub
Function P12Rx(S, R As RegExp) As P12
Dim M As Match: Set M = Mch(S, R): If IsNothing(M) Then Exit Function
With P12Rx
    .P1 = M.FirstIndex + 1
    .P2 = .P1 + M.Length
End With
End Function

Function HasPatn(S, Patn$) As Boolean:      HasPatn = HasRx(S, Rx(Patn)): End Function
Function HasRx(S, Rx As RegExp) As Boolean:   HasRx = Rx.Test(S):         End Function
Function HasRxAyOr(S, Rxay() As RegExp) As Boolean
If Si(Rxay) = 0 Then HasRxAyOr = True: Exit Function
Dim Rx: For Each Rx In Rxay
    If CvRx(Rx).Test(S) Then HasRxAyOr = True: Exit Function
Next
End Function
Function HasRxAyAnd(S, Rxay() As RegExp) As Boolean
If Si(Rxay) = 0 Then HasRxAyAnd = True: Exit Function
Dim Rx: For Each Rx In Rxay
    If Not CvRx(Rx).Test(S) Then Exit Function
Next
HasRxAyAnd = True
End Function

Function PatnSsubss$(Ssubss$)
Dim SySsub$(): SySsub = SySs(Ssubss)
PatnSsubss = Jn(SySsub, "|")
End Function

Function PosRx&(S, R As RegExp, Optional PosBeg% = 1)
Dim M As Match: Set M = Mch(Mid(S, PosBeg + 1), R)
If Not IsNothing(M) Then PosRx = M.FirstIndex + 1 + PosBeg
End Function
Function PosPatn&(S, Patn$, Optional PosBeg% = 1):   PosPatn = PosRx(S, Rx(Patn), PosBeg): End Function
Function PosQuoDbl(S, Optional PosBeg% = 1):       PosQuoDbl = InStr(PosBeg, S, vbQuoDbl): End Function
Private Sub B_Mch()
Dim A As Match
Dim R  As RegExp: Set R = Rx("m[ae]n")
Set A = Mch("amensdklf", R)
Ass A.FirstIndex = 1
Ass A.Length = 3
Stop
End Sub
Function Mchcoll(S, R As RegExp) As MatchCollection: Set Mchcoll = R.Execute(S): End Function 'MchColl:: #Match-Collection-Object#
Function Mch(S, R As RegExp) As Match
Dim Mc As MatchCollection: Set Mc = Mchcoll(S, R)
If Mc.Count = 0 Then Exit Function
Set Mch = Mc(0)
End Function
