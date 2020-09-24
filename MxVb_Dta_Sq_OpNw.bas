Attribute VB_Name = "MxVb_Dta_Sq_OpNw"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Sq_OpNw."
Function SqSsly(Ssly$()) As Variant(): SqSsly = SqDy(WDySsly(Ssly)): End Function
Private Function WDySsly(Ssly$()) As Variant()
Dim Ss: For Each Ss In Itr(Ssly)
    PushI WDySsly, SySs(Ss)
Next
End Function
Function SqDotly(Dotly$()) As Variant(): SqDotly = SqDy(WDy(Dotly)): End Function
Function SqTsy(Tsy$()) As Variant():       SqTsy = SqDy(DyTsy(Tsy)): End Function
Private Function WDy(Dotly$()) As Variant()
Dim DotLn: For Each DotLn In Itr(Dotly)
    PushI WDy, SplitDot(DotLn)
Next
End Function

Function SqDy(Dy(), Optional SkpNDr = 0) As Variant()
Dim O(), NR&, NC&
NR = Si(Dy)
NC = NDcDy(Dy)
ReDim O(1 To NR - SkpNDr, 1 To NC)
Dim R&: For R = 1 To NR
    Dim Dr: Dr = Dy(R - 1)
    SetSqr O, R, Dr
Next
SqDy = O
End Function
Function SqDyByt(Dy(), Optional SkpNDr = 0) As Variant()
Dim O(), NR&, NC&
NR = Si(Dy)
NC = NDcDy(Dy)
ReDim O(1 To NR - SkpNDr, 1 To NC)
Dim R&: For R = 1 To NR
    Dim Dr: Dr = Dy(R - 1)
    SetSqr O, R, Dr
Next
SqDyByt = O
End Function

Function SqRC(R&, C&) As Variant() 'Ret : a Sq(1 to @R, 1 to @C)
Dim O()
ReDim O(1 To R, 1 To C)
SqRC = O
End Function

Function sampSq() As Variant()
Const NR% = 10
Const NC% = 10
Dim O(), R%, C%
ReDim O(1 To NR, 1 To NC)
sampSq = O
For R = 1 To NR
    For C = 1 To NC
        O(R, C) = R * 1000 + C
    Next
Next
sampSq = O
End Function
Function SqRowAp(ParamArray Ap()) As Variant(): Dim Av(): Av = Ap: SqRowAp = SqRow(Av): End Function
Function SqColAp(ParamArray Ap()) As Variant(): Dim Av(): Av = Ap: SqColAp = SqCol(Av): End Function
Function SqRow(Ay) As Variant()
Dim N&: N = Si(Ay): If N = 0 Then Exit Function
Dim O(): ReDim O(1 To 1, 1 To N)
Dim J&, V: For Each V In Ay
    J = J + 1
    O(1, J) = V
Next
SqRow = O
End Function
Function SqcLines(Lines$) As Variant(): SqcLines = SqCol(SplitCrLf(Lines)): End Function
Function SqCol(Ay) As Variant()
Dim N&: N = Si(Ay): If N = 0 Then Exit Function
Dim O(): ReDim O(1 To N, 1 To 1)
Dim J&, V: For Each V In Ay
    J = J + 1
    O(J, 1) = V
Next
SqCol = O
End Function
