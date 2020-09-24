Attribute VB_Name = "MxVb_Ay_Op_AyIns"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_OpIns."

Function AyIns2Ele(Ay, E1, E2, Optional BefIx = 0): AyIns2Ele = AyInsAy(Ay, Array(E1, E2), BefIx): End Function

Private Sub B_AyInsAy()
GoSub T1
GoSub T2
Exit Sub
Dim Ay, AyIns, IxBef&, Cnt%
T1:
    Ay = Array(1, 2, 3)
    AyIns = Array("a", "b")
    IxBef = 2
    Ept = Array(1, 2, "a", "b", 3)
    GoTo Tst
T2:
    Ay = Array(3)
    IxBef = 0
    Ept = Array(1, 3)
    GoTo Tst
Tst:
    Act = AyInsAy(Ay, AyIns, IxBef)
    C
    Return
End Sub

Function SyInsSy(Sy$(), SyIns$(), Optional IxBef = 0) As String():                      SyInsSy = AyInsAy(Sy, SyIns, IxBef):  End Function
Function SyIns(Sy$(), Optional Ele$, Optional IxBef = 0, Optional Cnt = 1) As String():   SyIns = AyIns(Sy, Ele, IxBef, Cnt): End Function

Private Sub B_AyIns()
'GoSub T1
GoSub T2
Dim Ay, IxBef, Cnt, Ele
Exit Sub
T1:
    Ay = Array(1, 2, 3)
    IxBef = 2
    Cnt = 1
    Ele = "X"
    Ept = Array(1, 2, "X", 3)
    GoTo Tst
T2:
    Ay = Array(1, 2, 3)
    IxBef = 1
    Cnt = 3
    Ele = Empty
    Ept = Array(1, Empty, Empty, Empty, 2, 3)
    GoTo Tst
Tst:
    Act = AyIns(Ay, Ele, IxBef, Cnt)
    C
    Return
End Sub
Function AyIns(Ay, Optional Ele = Empty, Optional IxBef = 0, Optional Cnt = 1)
Dim AyToIns
    AyToIns = AyNw(Ay)
    ReDim Preserve AyToIns(Cnt - 1)
    If Not IsEmpty(Ele) Then
        Dim J%: For J = 0 To Cnt - 1
            AyToIns(J) = Ele
        Next
    End If
AyIns = X_AyInsAy(Ay, CLng(IxBef), AyToIns)
End Function
Function AyInsAy(Ay, AyIns, Optional IxBef = 0): AyInsAy = X_AyInsAy(Ay, CLng(IxBef), AyIns): End Function

Private Function X_AyInsAy(Ay, IxBef&, AyIns)
Const CSub$ = CMod & "X_AyInsAy"
Dim N&: N = Si(Ay)
If Not IsBet(IxBef, 0, N) Then Thw CSub, "Given @IxBef is out of 0 - Si(@Ay)", "IxBef Si(@Ay)", IxBef, N
Dim UAyIns&: UAyIns = UB(AyIns)
Dim UNew&: UNew = N + UAyIns
Dim O
    O = Ay
    ReDim Preserve O(UNew)
    'Move
    Dim IxTo&: For IxTo = UNew To IxBef + UAyIns Step -1
        Dim IxFm&: IxFm = IxTo - UAyIns - 1
        O(IxTo) = O(IxFm)
    Next
    'Ins
    For IxTo = IxBef To IxBef + UAyIns
        O(IxTo) = AyIns(IxTo - IxBef)
    Next
X_AyInsAy = O
End Function
