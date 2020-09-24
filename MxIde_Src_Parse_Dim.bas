Attribute VB_Name = "MxIde_Src_Parse_Dim"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Parse_Dim."
Private Sub B_StmtyDimPC():                                  Brw StmtyDimPC:             End Sub
Function StmtyDimPC() As String():              StmtyDimPC = StmtyDimP(CPj):             End Function
Function StmtyDimP(P As VBProject) As String():  StmtyDimP = StmtyDim(Contlny(SrcP(P))): End Function
Private Sub B_StmtyDim():                                    Brw StmtyDim(SrcP(CPj)):    End Sub
Function StmtyDim(Src$()) As String(): Dim L: For Each L In Itr(Src): PushNB StmtyDim, StmtDim(L): Next: End Function
Function ItmyDim(Src$()) As String(): ItmyDim = ItmyDimStmty(StmtyDim(Src)): End Function
Function ItmyDimStmty(StmtyDim$()) As String()
Const CSub$ = CMod & "ItmyDimStmty"
Dim Stmt: For Each Stmt In Itr(StmtyDim)
    Dim L$: L = Stmt
    If Not IsShfPfxSpc(L, "Dim") Then Thw CSub, "Given StmtDim does not have Pfx [Dim ]", "Stmt", Stmt
    PushIAy ItmyDimStmty, W1ItmyAftDim(L)
Next
End Function
Private Sub B_W1ItmyAftDim()
Dim AftDim
GoSub T1
T1:
    AftDim = "A(1, 2), B(), C(1, 2)"
    Ept = Sy("A(1, 2)", "B()", "C(1, 2)")
    GoTo Tst
Tst:
    Act = W1ItmyAftDim(AftDim)
    C
    Return
End Sub
Private Function W1ItmyAftDim(AftDim) As String()
W1ItmyAftDim = AmTrim(SyPosy(AftDim, W1Posy(AftDim)))
End Function
Private Function W1Posy(AftDim) As Integer()
Dim P%(): P = PosySsub(AftDim, ",")
Dim I: For Each I In Itr(P)
    If W1IsPosCmaNotInBkt(AftDim, I) Then
        PushI W1Posy, I
    End If
Next
End Function
Private Function W1IsPosCmaNotInBkt(AftDim, PosCma) As Boolean
Dim IsBktLeft As Boolean: IsBktLeft = W1IsBktFstOnLeft(AftDim, PosCma)
Dim IsBktRight As Boolean: IsBktRight = W1IsBktFstOnRight(AftDim, PosCma)
Select Case True
Case IsBktLeft And IsBktRight: Exit Function
Case Not IsBktLeft And Not IsBktRight: W1IsPosCmaNotInBkt = True: Exit Function
End Select
Thw CSub, "@PosCma does not have Bkt on Left and Right either both true or both false", "@PosCma @AftDim IsBktLeft IsBktRight", PosCma, AftDim, IsBktLeft, IsBktRight
End Function
Private Function W1IsBktFstOnLeft(AftDim, PosCma) As Boolean
Dim J%: For J = PosCma - 1 To 1 Step -1
    Select Case Mid(AftDim, J, 1)
    Case ",": Exit Function
    Case "(": W1IsBktFstOnLeft = True: Exit Function
    End Select
Next
End Function
Private Function W1IsBktFstOnRight(AftDim, PosCma) As Boolean
Dim J%: For J = PosCma + 1 To Len(AftDim)
    Select Case Mid(AftDim, J, 1)
    Case ",": Exit Function
    Case ")": W1IsBktFstOnRight = True: Exit Function
    End Select
Next
End Function
Function StmtDim$(Contln)
Dim Stmt: For Each Stmt In Itr(Stmty(Contln))
    If HasPfxSpc(Contln, "Dim") Then StmtDim = Stmt: Exit Function
Next
End Function
