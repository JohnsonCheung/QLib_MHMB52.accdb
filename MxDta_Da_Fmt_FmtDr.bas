Attribute VB_Name = "MxDta_Da_Fmt_FmtDr"
Option Compare Database
Option Explicit
Private Sub B_W2SqDr()
GoSub T1
Exit Sub
Dim Dr, W%(), Zer As eZer, BoolyAliR() As Boolean, Act$(), Ept$(), URow%
T1:
    Dr = Array(123#, "A", #12/31/2020#, RplVBar("a|b|c"))
    Ept = Array(): ReDim Ept(1 To 3, 1 To 5)
    URow = 1
    SetSqr Ept, 1, Array(12, 123#, "A", #12/31/2020#, "a")
    SetSqr Ept, 2, Array(, , , , "b")
    SetSqr Ept, 3, Array(, , , , "c")
    GoTo Tst
Tst:
    Act = W2SqDr(Dr, W, Zer, BoolyAliR)
    Ass IsEqSq(Ept, Act)
    Stop
End Sub

Function FmtDr(Dr, Optional F As eTblFmt = eTblFmtSS) As String(): FmtDr = FmtDrPm(Dr, WdtyDr(Dr), eZerShw, QmkDta(F)): End Function
Function FmtDrPm(Dr, W%(), Z As eZer, Q As Qmk, Optional BoolyAliR) As String()   ' Return @@Dy from @Dr.  Because @Dr-ele may have multi-lines.  If no ele in @Dr
Dim Sq$(): Sq = W2SqDr(Dr, W, Z, BoolyAliR)
Dim J%: For J = 1 To NDrSq(Sq)
    Dim DrI(): DrI = DrSq(Sq, J)
    PushI FmtDrPm, LnDr(DrI, Q)
Next
End Function
Private Function W2SqDr(Dr, W%(), Zer As eZer, BoolyAliR) As String()   'ret a @@Sq from @Dr due to ele in @Dr may be multiple line and each Dc of @@Sq is ali by @W
If Si(Dr) = 0 Then Exit Function
Dim Dcy()
    If IsBooly(BoolyAliR) Then
        Dcy = W2DcyDrAliYes(Dr, W, Zer, CvBooly(BoolyAliR))
    Else
        Dcy = W2DcyDrAliNo(Dr, W, Zer)
    End If
Dim OSq$()
    Dim NR%: NR = W2NRowDcy(Dcy)
    Dim NC%: NC = Si(Dr)
    ReDim OSq(1 To NR, 1 To NC)
'Set OSq
    Dim Cno%: Cno = 0
    Dim DcI: For Each DcI In Dcy
        Cno = Cno + 1
        W2SetSqDc OSq, Cno, CvSy(DcI), W(Cno - 1) ' Set Each *OSq with same width
        Dim Rno%: For Rno = 1 To NR
            If Len(OSq(Rno, Cno)) <> W(Cno - 1) Then Stop
        Next
    Next
W2SqDr = OSq
End Function
Private Function W2DcyDrAliNo(Dr, W%(), Zer As eZer) As Variant()
Dim V, J%: For Each V In Dr
    PushI W2DcyDrAliNo, W2DcV(V, W(J), Zer)
    J = J + 1
Next
End Function
Private Function W2DcyDrAliYes(Dr, W%(), Zer As eZer, BoolyAliR() As Boolean) As Variant()
Dim V, J%: For Each V In Dr
    PushI W2DcyDrAliYes, W2DcV(V, W(J), Zer, BoolyAliR(J))
    J = J + 1
Next
End Function
Private Function W2DcV(V, W%, Zer As eZer, Optional IsAliR As Boolean) As String()    ' Ret a Dc from V with each Ln limited by @W and handle @Zer
Select Case True
Case IsDte(V):           W2DcV = Sy(Ali(StrDte(CDate(V)), W, IsAliR))
Case W2IsHidZer(V, Zer):
Case IsStr(V):           W2DcV = W2DcStrCell(V, W, IsAliR)
Case Else:               W2DcV = Sy(Ali(V, W, IsAliR))
End Select
End Function
Private Function W2DcStrCell(Cell, W%, IsAliR As Boolean) As String()
Select Case True
Case IsLines(Cell)
    W2DcStrCell = WrdlnyLines(Cell, W)
Case Len(Cell) > W
    W2DcStrCell = WrdlnyLines(Cell, W)
Case Else
    W2DcStrCell = Sy(Ali(Cell, W, IsAliR))
End Select
End Function

Private Function W2NRowDcy%(Dcy())
Dim O%: O = 1
Dim DcI: For Each DcI In Dcy
    O = Max(O, Si(DcI))
Next
W2NRowDcy = O
End Function
Private Sub W2SetSqDc(OSq$(), Cno%, Dc$(), W%)
Dim U%: U = UB(Dc)
Dim IxR%, L: For IxR = 0 To U
    Dim V$: V = Dc(IxR)
    OSq(IxR + 1, Cno) = V & Space(W - Len(V))
Next
Dim S$: S = Space(W)
For IxR = U + 1 To UBound(OSq, 1) - 1
    OSq(IxR + 1, Cno) = S
Next
End Sub
Private Function W2IsHidZer(V, Zer As eZer) As Boolean
Select Case True
Case _
    Zer <> eZerHid, _
    IsNumeric(V), _
    V <> 0
Case Else
    W2IsHidZer = True
End Select
End Function


