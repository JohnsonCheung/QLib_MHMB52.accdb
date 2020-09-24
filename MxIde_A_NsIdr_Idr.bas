Attribute VB_Name = "MxIde_A_NsIdr_Idr"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_A_NsIdr_Idr."
Private Sub B_RplNonAlphNum():                          Brw NonAlphaNum_Rpl(SrclPC): End Sub
Function NonAlphaNum_Is(S) As Boolean: NonAlphaNum_Is = NonAlphaNum_Rx.Test(S):      End Function
Function NonAlphaNum_Rx() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("/[^_0-9a-zA-Z.\n\r]/gm")
Set NonAlphaNum_Rx = X
End Function
Function NonAlphaNum_Rpl$(S): NonAlphaNum_Rpl = NonAlphaNum_Rx.Replace(S, " "): End Function

Private Sub B_NyS()
Dim S$
Dim Text
GoSub Z
'GoSub T0
Exit Sub
Z:
    Dim Lines$: Lines = SrclPC
    Dim Ny1$(): Ny1 = NyStr(Lines)
    Dim Ny2$(): Ny2 = Idry(Lines)
    If Not IsEqAy(Ny1, Ny2) Then Stop
    Return
T0:
    S = "S_S"
    Ept = Sy("S_S")
    GoTo Tst
Tst:
    Act = NyStr(S)
    C
    Return
End Sub

Private Sub B_NmAet():                       VcAet AetSrt(NmAet(SrclPC)):         End Sub
Function NmAet(S) As Dictionary: Set NmAet = AetAy(NyStr(S)):                     End Function
Function NyStr(S) As String():       NyStr = AwNm(SySs(RplLf(RplCr(RplPun(S))))): End Function

Private Sub B_DiIdrqCnt()
'There will be 3 subMch for these patn (()|()): SubMch1 is the outer bkt and SubMch2 and 3 are the inner.  If SubMch2 then 3 will be empty, of SubMch3, 2 will be empty.
Dim A As Dictionary: Set A = DiIdrqCnt(JnCrLf(SrcP(CPj)))
Set A = DiSrt(A)
BrwDi A
End Sub
Private Sub B_IdrStsLines()
Debug.Print IdrStsLy(SrcP(CPj))
End Sub

Function IdrStsLines$(Lines)
IdrStsLines = IdrStsLy(SplitCrLf(Lines))
End Function

Sub CntIdrP()
Debug.Print IdrStsLy(SrcP(CPj))
End Sub

Function IdrStsLy$(Ly$())
Dim W&, D&, Sy$(), B&, L&, S$
S = JnCrLf(Ly)
Sy = Idry(S)
W = Si(Sy)
D = Si(AwDis(Sy))
L = Si(Ly)
B = Len(S)
IdrStsLy = IdrSts(B, L, W, D)
End Function

Function IdrSts$(B&, L&, W&, D&)
Dim BB As String * 9: RSet BB = B
Dim LL As String * 9: RSet LL = L
Dim X_ As String * 9: RSet X_ = W
Dim DD As String * 9: RSet DD = D
IdrSts = FmtQQ("Len            : ?|Lines          : ?|Words          : ?|Distinct Words : ?", BB, LL, X_, DD)
End Function

Function NIdr&(S)
NIdr = Si(Idry(S))
End Function

Function NDistIdr&(S)
NDistIdr = Si(AwDis(Idry(S)))
End Function

Function DiIdrqCnt(S) As Dictionary
Set DiIdrqCnt = DiCnt(Idry(S))
End Function

Function AetIdr(S) As Dictionary
Set AetIdr = AetAy(Idry(S))
End Function

Function FstIdrAetP() As Dictionary
Set FstIdrAetP = New Dictionary
Dim L: For Each L In Itr(SrcRmvVmk(SrcP(CPj)))
    PushAetEle FstIdrAetP, IDr(L)
Next
End Function

Private Sub B_Idr()
Dim S$
GoSub T1
Exit Sub
T1:
    S = "00A cB"
    Ept = "cB"
    GoTo Tst
Tst:
    Act = IDr(S)
    C
    Return
End Sub

Private Sub B_Idry()
Dim S
'GoSub T1
GoSub Z1
'GoSub Z2
Z1:
    Dim O$(): For Each S In SrcPC
        PushI O, S & vbCrLf & JnSpc(Idry(S))
    Next
    BrwLsy O
    Return
Z2:
    VcAy AySrtQ(AwDis(Idry(SrclPC)))
    Return
T1:
    S = "Function 0AA B"
    Ept = Sy("Function", "B")
    GoTo Tst
Tst:
    Act = Idry(S)
    C
    Return
End Sub
Function IDr$(S): IDr = SsubRx(S, WRx): End Function
Function Idry(S) As String() ' Identifier array of @S
Dim M As Match: For Each M In WRx.Execute(S)
    Stop 'PushI Idry, ZZIdrMch(M)
Next
End Function
Private Function WRx() As RegExp
Const P$ = "/(^[A-Z]\w*)|[ .\(]([A-Z]\w*)/mi" ' Rx should use ignorecas.  2 cases: a word is begin of a line or fst chr is one of these char [ .(]..
Dim X As RegExp: If IsNothing(X) Then Set X = Rx(P)
Set WRx = X
End Function

Function NNmChrRx() As RegExp ' Non name char regexp
Dim O As New RegExp
Set O = Rx("\W")
End Function

Function RplNNmChr$(S) ' replace non name char to space
'NNmChr:Cml #non-nm-chr#
RplNNmChr = NNmChrRx.Replace(S, " ")
End Function

Function AmRplNNmChr(Ay) As String()
Dim L: For Each L In Itr(Ay)
    PushI AmRplNNmChr, RplNNmChr(L)
Next
End Function

Private Sub B_AmRplNNmChr()
Brw AmRplNNmChr(SrcP(CPj))
End Sub

Function IdrySrc(Src$()) As String()
Stop 'Dim L: For Each L In Itr(SrcRplVmk(Src))
    'PushI IdrySrc, Idry(L)
'Next
End Function

Function HasIdr(S, IDr) As Boolean
HasIdr = HasEle(Idry(Sy(S)), IDr)
End Function

Function Idrss$(S)
Idrss = JnSpc(Idry(S))
End Function

Function IdrssAy(Sy$()) As String()
Dim S: For Each S In Itr(Sy)
    PushI IdrssAy, Idrss(S)
Next
End Function


Function IdrLblLinPos$(IdrPos%(), OFmNo&)
Dim O$(), A$, B$, W%, J%
If Si(IdrPos) = 0 Then Exit Function
PushNB O, Space(IdrPos(0) - 1)
For J = 0 To UB(IdrPos) - 1
    A = OFmNo
    W = IdrPos(J + 1) - IdrPos(J)
    If W > Len(A) Then
        A = AliL(A, W)
        If Len(A) <> W Then Stop
    Else
        A = Space(W)
    End If
    PushI O, A
    OFmNo = OFmNo + 1
Next
A = OFmNo
PushI O, A
IdrLblLinPos = Jn(O)
End Function
Function IdrLblLin(Ln, OFmNo&)
IdrLblLin = IdrLblLinPos(IdrPosy(Ln), OFmNo)
End Function

Function IdrPosy(Ln) As Integer()
Dim J%, LasIsSpc As Boolean, CurIsSpc As Boolean
LasIsSpc = True
For J = 1 To Len(Ln)
    CurIsSpc = Mid(Ln, J, 1) = " "
    Select Case True
    Case CurIsSpc And LasIsSpc
    Case CurIsSpc:          LasIsSpc = True
    Case LasIsSpc:          PushI IdrPosy, J
                            LasIsSpc = False
    Case Else
    End Select
Next
End Function
Function IdrLblLinPairLno(Ln, Lno, LnoWdt, OFmNo&) As String()
Dim O$(): O = IdrLblLinPair(Ln, OFmNo)
O(0) = Space(LnoWdt) & " : " & O(0)
'O(1) = AliL(Lno, LnoWdt) & " : " & O(1)
IdrLblLinPairLno = O
End Function
Function IdrLblLinPair(Ln, OFmNo&) As String()
PushI IdrLblLinPair, IdrLblLin(Ln, OFmNo)
PushI IdrLblLinPair, Ln
End Function
Function IdrLblLy(Ly$(), OFmNo&) As String()
Dim J&, LnoWdt%, A$
A = UB(Ly)
LnoWdt = Len(A)
For J = 1 To UB(Ly)
    PushIAy IdrLblLy, IdrLblLinPairLno(Ly(J), J, LnoWdt, OFmNo)
Next
End Function


Private Sub B_IdrLblLin()
Dim Ln, FmNo&
GoSub T0
Exit Sub
T0:
    FmNo = 2
    '               10        20        30        40        50        60
    '      123456789 123456789 123456789 123456789 123456789 123456789 123456789
    Ln = "Lbl01 Lbl02 Lbl03    Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10"
    Ept = "2     3     4        5     6     7     8     9     10    11"
    GoTo Tst
Tst:
    Act = IdrLblLin(Ln, FmNo)
    C
    Return
End Sub
Private Sub B_IdrPosy()
Dim Ln
GoSub T0
Exit Sub
T0:
    '               10        20        30        40        50        60
    '      123456789 123456789 123456789 123456789 123456789 123456789 123456789
    Ln = "Lbl01 Lbl02 Lbl03    Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10"
    Ept = Inty(1, 7, 13, 22, 28, 34, 40, 46, 52, 58)
    GoTo Tst
Tst:
    Act = IdrPosy(Ln)
    C
    Return
End Sub

Private Sub B_IdrLblLy()
Dim Fm&: Fm = 1
Brw IdrLblLy(SrcP(CPj), Fm)
End Sub

Function IdryPC() As String(): IdryPC = IdryP(CPj): End Function
Function IdryP(P As VBProject) As String()
Dim L: For Each L In SrcP(P)
    PushIAy IdryP, Idry(L)
Next
End Function
'== ZZ
Private Function ZZIdrzMch$(M As Match)
Const CSub$ = CMod & "ZZIdrzMch"
Dim S As ISubMatches: Set S = M.SubMatches
If S.Count <> 2 Then ThwImposs CSub, "The ZZIdrPatn should be ()|() so that it will gives 2 subMatch, but now the submatch count=[" & S.Count & "]"
If IsEmpty(S(1)) Then
    ZZIdrzMch = S(0) ' fst-SMch match means the idr started with begin of a line
Else
    ZZIdrzMch = S(1)  ' snd-SMch match means the idr started with 1 of spc of (
End If
End Function
