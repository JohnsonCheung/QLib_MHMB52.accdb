Attribute VB_Name = "MxVb_NNy"
'NNy:Cml :Ly #NN-Array#
'NN:Cml  :Ln #Name-Name# each name is separated only by one space
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_NNy."

Private Sub B_SwapParChdNNy()
Const P$ = "MxVbNNy\SwapParChdNNy\"
Const Cas1NNyFnSeg$ = P & "Cas1-NNy.txt"
Const Cas1EptFnSeg$ = P & "Cas1-Ept.txt"
Dim NNy$()
T1:
    EdtRES Cas1NNyFnSeg
    EdtRES Cas1EptFnSeg
    Stop
    NNy = Resy(Cas1NNyFnSeg)
    Ept = Resy(Cas1EptFnSeg)
    GoTo Tst
Tst:
    Act = SwapParChdNNy(NNy)
    C
    Return
End Sub
Function SwapParChdNNy(NNy$()) As String() ' return a new NNy with parent and children are swap
'Parent becomes child and child becomes child
'First Nm is a parent
Dim P$(), C$()
Dim NN: For Each NN In Itr(NNy)
    WAsgParChd NN, P, C
Next
Dim ChdIxGpy() As IxGp: ChdIxGpy = SamEleIxGp(C)
Dim J&: For J = 0 To IxGpUB(ChdIxGpy)
    PushI SwapParChdNNy, WSwappedln(C, ChdIxGpy(J).Ixy, P)
Next
End Function
Private Sub WAsgParChd(NN, OPar$(), OChd$())
Dim N$(): N = SplitSpc(NN)
Dim P$: P = N(0)
PushI OPar, P
PushI OChd, N(1)
Dim J%: For J = 2 To UB(N)
    PushI OPar, P
    PushI OChd, N(J)
Next
End Sub
Private Function WSwappedln$(Chd$(), Ixy&(), Par$())
Const CSub$ = CMod & "WSwappedln"
Dim C$: C = Chd(Ixy(0))
Dim J%
If True Then
    For J = 1 To UB(C)
        If C <> Chd(Ixy(J)) Then
            Thw CSub, "Given Ixy should pointing to @Chd with same value", "Fst-Chd [J which has dif Chd] Chd", C, J, Chd
        End If
    Next
End If
Dim P$(): For J = 0 To UB(Ixy)
    PushI P, Par(Ixy(J))
Next
WSwappedln = C & " " & JnSpc(P)
End Function

Private Sub B_FmtNNy()
Dim NNy$(), NCol%
GoSub T1
Exit Sub
T1:
    Erase NNy
    PushI NNy, "A1 B1 C1"
    PushI NNy, "A2 B2 C2 D2"
    PushI NNy, "A3 B3 C3 D3 E3"
    PushI NNy, "A4 B4 C4 D4 E4 F4"
    PushI NNy, "A5 B5 C5 D5 E5 F5 G5"
    PushI NNy, "A6 B6 C6 D6 E6 F6 G6 H6"
    PushI NNy, "A7 B7 C7 D7 E7 F7 G7 H7 I7"
    
    NCol = 4
    Ept = SyEmp
    PushI Ept, "A1 B1 C1"
    PushI Ept, "A2 B2 C2 D2"
    PushI Ept, "A3 B3 C3 D3"
    PushI Ept, ".  E3"
    PushI Ept, "A4 B4 C4 D4"
    PushI Ept, ".  E4 F4"
    PushI Ept, "A5 B5 C5 D5"
    PushI Ept, ".  E5 F5 G5"
    PushI Ept, "A6 B6 C6 D6"
    PushI Ept, ".  E6 F6 G6"
    PushI Ept, ".  H6"
    PushI Ept, "A7 B7 C7 D7"
    PushI Ept, ".  E7 F7 G7"
    PushI Ept, ".  H7 I7"
    GoTo Tst
Tst:
    Act = FmtNNy(NNy, NCol)
    C
    Return
End Sub
Function FmtNNy(NNy$(), Optional NCol% = 11) As String()
Dim Dy()
Dim NN: For Each NN In Itr(NNy)
    PushIAy Dy, W2WrpNN(NN, NCol)
Next
'FmtNNy = AliStrcUy(StrcUyzDy(Dy))
End Function
Private Function W2WrpNN(NN, NCol%) As String()
Dim Ny$(): Ny = SplitSpc(NN)
Dim UBlk%: UBlk = W2UBlk(Si(Ny), NCol)
Dim J%: For J = 0 To UBlk
'    PushI W2WrpNN, W2Ln(N, J, NCol)
Next
End Function
Private Sub B_W2UBlk()
Dim NNm&, NCol%
GoSub T1
Exit Sub
T1:
    NCol = 10
    Ept = 1
    For NNm = 0 To 10
        GoSub Tst
    Next
    '
    Ept = 2
    For NNm = 11 To 19
        GoSub Tst
    Next
    '
    Ept = 3
    For NNm = 20 To 28
        GoSub Tst
    Next
    Return
Tst:
    Act = W2UBlk(NNm, NCol)
    C
    Return
End Sub
Private Function W2UBlk%(NNm&, NCol%)
Select Case True
Case NNm <= NCol: W2UBlk = 1
Case Else
    W2UBlk = ((NNm - NCol) \ NCol) + 2
End Select
End Function
Private Function W2Ln$(Ny$(), IBlk%, NCol%)
Dim B%: B = W2Bix(IBlk, NCol)
Dim E%: E = W2Eix(B, IBlk, NCol)
W2Ln = IIf(IBlk = 0, "", ". ") & JnSpc(AwBE(Ny, B, E))
End Function
Private Function W2Bix%(IBlk%, NCol%)
If IBlk = 0 Then
    W2Bix = 0
Else
    W2Bix = NCol + IBlk * (NCol - 1)
End If
End Function
Private Function W2Eix%(Bix%, IBlk%, NCol%)
If IBlk = 0 Then
    W2Eix = NCol - 1
Else
    W2Eix = Bix + NCol - 1
End If
End Function
