Attribute VB_Name = "MxIde_Src_Vmk_Brk"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Vmk_Brk."

Private Sub B_PosVmk()
Dim S$(): S = SrcPC
Dim Posy%(): Posy = W1Posy(S)
VcAy W1LyVmkNo(S, Posy%()), "Vmk-No "
VcAy W1LyVmkYes(S, Posy), "Vmk-Yes "
End Sub
Private Function W1Posy(Src$()) As Integer()
Dim L: For Each L In Src
    PushI W1Posy, PosVmk(L)
Next
End Function
Private Function W1LyVmkNo(Src$(), Posy%()) As String()
Dim J&, P: For Each P In Posy
    If P = 0 Then PushI W1LyVmkNo, Src(J)
    J = J + 1
Next
End Function
Private Function W1LyVmkYes(Src$(), Posy%()) As String()
Dim J&, P: For Each P In Posy
    If P > 0 Then PushIAy W1LyVmkYes, W1Ly3Ln(Src(J), P)
    J = J + 1
Next
End Function
Private Function W1Ly3Ln(L, Pos) As String()
PushI W1Ly3Ln, L
PushI W1Ly3Ln, StrDup(" ", Pos - 1) & "^"
PushI W1Ly3Ln, ""
End Function
Private Sub B_BrkVmk()
GoSub T1
Exit Sub
Dim Contln, Act As S12, Ept As S12
T1:
    Contln = "'A"
    Ept = S12("", "A")
    GoTo Tst
Tst:
    Act = BrkVmk(Contln)
    Ass IsEqS12(Act, Ept)
    Return
End Sub
Sub AsgBrkVmk(Contln, OBefVmk$, OVmk$)
With BrkVmk(Contln)
    OBefVmk = .S1
    If .S2 = "" Then
        OVmk = ""
    Else
        OVmk = "' " & .S2
    End If
End With
End Sub
Function BrkVmk(Contln) As S12: BrkVmk = Brk1At(Contln, PosVmk(Contln), NoTrim:=True): End Function
Private Function PosVmk%(Ln)
Dim P%: P = InStr(Ln, "'"): If P = 0 Then Exit Function
If Trim(Left(Ln, P - 1)) = "" Then PosVmk = P: Exit Function
Dim Py%(): Py = PosySsub(Ln, "'")
Dim J: For Each J In Py
    If Not WIsPosInDblQ(Ln, J) Then PosVmk = J: Exit Function
Next
End Function
Private Function WIsPosInDblQ(Ln, Pos) As Boolean
Dim Bef$: Bef = Left(Ln, Pos - 1)
WIsPosInDblQ = IsOdd(NSsub(Bef, vbQuoDbl))
End Function
Function Vmk$(Ln): Vmk = TakPosIf(Ln, PosVmk(Ln)): End Function

Function RmvVmk$(Ln)
Dim P%: P = PosVmk(Ln): If P = 0 Then RmvVmk = Ln: Exit Function
RmvVmk = RTrim(Left(Ln, P - 1))
End Function
