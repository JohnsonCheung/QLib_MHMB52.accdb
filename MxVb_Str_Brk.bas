Attribute VB_Name = "MxVb_Str_Brk"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Brk."

Function BrkTRst(S) As S12
Dim T$, Rst$
AsgT1r S, T, Rst
BrkTRst = S12(T, Rst)
End Function



Function Brk(S, Sep$, Optional NoTrim As Boolean) As S12
Const CSub$ = CMod & "Brk"
Dim P&: P = InStr(S, Sep)
If P = 0 Then Thw CSub, "{S} does not contains {Sep}", "S Sep", S, "[" & Sep & "]"
Brk = BrkPos(S, P, Sep, NoTrim)
End Function
Function Brk1(S, Sep$, Optional NoTrim As Boolean) As S12
Dim P&: P = InStr(S, Sep)
If P = 0 Then Brk1 = S12Nw(S, "", NoTrim): Exit Function
Brk1 = Brk1Pos(S, P, Sep, NoTrim)
End Function

Function Brk1Rev(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStrRev(S, Sep)
If P = 0 Then Brk1Rev = S12Nw(S, "", NoTrim): Exit Function
Brk1Rev = Brk1Pos(S, P, Sep, NoTrim)
End Function

Function Brk2(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStr(S, Sep)
Brk2 = Brk2Pos(S, P, Sep, NoTrim)
End Function

Function Brk2Pos(S, P&, Sep, NoTrim As Boolean) As S12
If P = 0 Then
    If NoTrim Then
        Brk2Pos = S12("", S)
    Else
        Brk2Pos = S12("", Trim(S))
    End If
    Exit Function
End If
Brk2Pos = Brk1Pos(S, P, Sep, NoTrim)
End Function

Function Brk2Rev(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStrRev(S, Sep)
Brk2Rev = Brk2Pos(S, P, Sep, NoTrim)
End Function

Function BrkPos(S, P&, Sep, NoTrim As Boolean) As S12
Dim S1$, S2$
S1 = Left(S, P - 1)
S2 = Mid(S, P + Len(Sep))
BrkPos = S12Nw(S1, S2, NoTrim)
End Function
Function Brk1At(S, At&, Optional NoTrim As Boolean) As S12
If At = 0 Then
    Brk1At = S12(S, "")
Else
    Brk1At = S12(Left(S, At - 1), Mid(S, At))
End If
If Not NoTrim Then Brk1At = S12Trim(Brk1At)
End Function

Function Brk1Pos(S, P&, Sep, Optional NoTrim As Boolean) As S12
If P = 0 Then
    Brk1Pos = S12Nw(S, "", NoTrim)
Else
    Brk1Pos = BrkPos(S, P, Sep, NoTrim)
End If
End Function


Function BrkBoth(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    BrkBoth = S12Nw(S, S, NoTrim)
    Exit Function
End If
BrkBoth = Brk1Pos(S, P, Sep, NoTrim)
End Function

Function BrkRev(S, Sep, Optional NoTrim As Boolean) As S12
Dim P&: P = InStrRev(S, Sep)
If P = 0 Then Raise "BrkRev: Str[" & S & "] does not contains Sep[" & Sep & "]"
BrkRev = Brk1Pos(S, P, Len(Sep), NoTrim)
End Function

Private Sub B_Brk1Rev()
Dim S1$, S2$, ExpS1$, ExpS2$, S
S = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(S, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub
