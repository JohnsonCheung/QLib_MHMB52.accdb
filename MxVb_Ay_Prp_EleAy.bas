Attribute VB_Name = "MxVb_Ay_Prp_EleAy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Ele."
Function EleFst(Ay)
If Si(Ay) = 0 Then Exit Function
Asg Ay(0), _
    EleFst
End Function

Function EleLas(Ay)
Dim N&: N = Si(Ay)
If N = 0 Then Exit Function
Asg Ay(N - 1), _
    EleLas
End Function

Function EleMin(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = EleFst(Ay)
Dim I: For Each I In Itr(Ay)
    If I < O Then O = I
Next
EleMin = O
End Function

Function EleMax(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = EleFst(Ay)
Dim I: For Each I In Itr(Ay)
    If I > O Then O = I
Next
EleMax = O
End Function

Function EleLasSnd(Ay)
Const CSub$ = CMod & "EleLasSnd"
Dim N&: N = Si(Ay)
If N <= 1 Then
    Thw CSub, "Only 1 or no ele in Ay"
Else
    Asg Ay(N - 2), EleLasSnd
End If
End Function
Private Function WFunElePmErMsg$(Ix, Ay)
WFunElePmErMsg = FmtQQ("Ix-? must between 0 to ? of an array-of-type-[?]", , Ix, UB(Ay), TypeName(Ay))
End Function
Function EleMust(Ay, Ix)
If IsBet(Ix, 0, UB(Ay)) Then
    EleMust = Ay(Ix)
Else
    ThwPm CSub, WFunElePmErMsg(Ix, Ay)
End If
End Function
Function EleMsg(Ay, Ix)
If IsBet(Ix, 0, UB(Ay)) Then
    EleMsg = Ay(Ix)
Else
    EleMsg = WFunElePmErMsg(Ix, Ay)
End If
End Function
Function Ele(Ay, Ix)
If IsBet(Ix, 0, UB(Ay)) Then Ele = Ay(Ix)
End Function
Function ElePfx$(Sy$(), Pfx$, Optional C As eCas)
Dim J%: For J = 0 To UB(Sy)
    If HasPfx(Sy(J), Pfx, C) Then
        If J = UB(Sy) Then
            ElePfx = Sy(0)
        Else
            ElePfx = Sy(J + 1)
        End If
        Exit Function
    End If
Next
ElePfx = Sy(0)
End Function
Function CixDrs%(D As Drs, C$): CixDrs = Cix(D.Fny, C): End Function



Function EleFstXP(Ay, XP$, P$)
Dim X: For Each X In Ay
    If Run(XP, X, P) Then
        Asg EleFstXP, _
            X
        Exit Function
    End If
Next
End Function


