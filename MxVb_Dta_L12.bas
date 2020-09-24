Attribute VB_Name = "MxVb_Dta_L12"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_L12."
Type L12: L1 As Long: L2 As Long: End Type 'Deriving(Ay) Fmt(CCml)
Function L12(L1, L2) As L12
If L1 < 0 Then Thw CSub, "L1 cannot <0", "L1 L2", L1, L2
If L2 < 0 Then Thw CSub, "L2 cannot <0", "L1 L2", L1, L2
If L2 < L1 Then Thw CSub, "L2 cannot < L1", "L1 L2", L1, L2
With L12
    .L1 = L1
    .L2 = L2
End With
End Function
Sub PushL12(O() As L12, M As L12): Dim N&: N = SiL12(O): ReDim Preserve O(N): O(N) = M: End Sub
Function UbL12&(A() As L12): UbL12 = SiL12(A) - 1: End Function
Function SiL12&(A() As L12): On Error Resume Next: SiL12 = UBound(A) + 1: End Function
Function LinesL12y$(A() As L12): LinesL12y = JnCrLf(LyL12y(A)): End Function
Function LyL12y(A() As L12) As String()
Dim J&: For J = 0 To UbL12(A)
    PushI LyL12y, LnL12(A(J))
Next
End Function
Function LnL12$(A As L12): LnL12 = "L12 " & A.L1 & " " & A.L2: End Function
Private Sub B_L12yNumy()
GoSub T1
Exit Sub
Dim NumyOrdered&()
T1:
    NumyOrdered = Lngy(1, 2, 3, 4, 6, 10)
    Ept = RplVBar("L12 1 4|L12 6 6|L12 10 10")
    GoTo Tst
Tst:
    Act = LinesL12y(L12yNumy(NumyOrdered, CSub))
    C
    Return
End Sub
Function L12yNumy(NumyOrdered, Optional Fun$ = "L12yNumy") As L12()
Dim O() As L12
    Dim Cur&, M As L12
    M.L1 = NumyOrdered(0)
    M.L2 = NumyOrdered(0)
    Dim Ix&: For Ix = 1 To UB(NumyOrdered)
        Cur = NumyOrdered(Ix)
        Select Case Cur - M.L2
        Case 0: Thw Fun, "There is DupNum in NumyOrdered", "CurNumIx CurNumDup Cur-M NumyOrdered", Ix, Cur, LnL12(M), AmAddIxPfx(NumyOrdered, 0)
        Case Is < 0: Thw CSub, "There Num not in order", "CurNumIx CurNum Cur-M NumyOrdered", Ix, Cur, LnL12(M), AmAddIxPfx(NumyOrdered, 0)
        Case 1: M.L2 = Cur
        Case Is > 1
            PushL12 O, M
            M = L12(Cur, Cur)
        Case Else: ThwImposs CSub, "Case should be reached here"
        End Select
    Next
    PushL12 O, M
L12yNumy = O
End Function
