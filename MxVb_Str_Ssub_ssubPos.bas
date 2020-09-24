Attribute VB_Name = "MxVb_Str_Ssub_ssubPos"
Option Compare Text
Option Explicit

Function SyPosy(S, Posy) As String()
Dim U%: U = UB(Posy): If U = -1 Then PushI SyPosy, S: Exit Function
PushI SyPosy, Left(S, Posy(0) - 1)
Dim J%: For J = 0 To UB(Posy)
    If J = U Then
        PushI SyPosy, Mid(S, Posy(J) + 1)
    Else
        PushI SyPosy, Mid(S, Posy(J) + 1, Posy(J + 1) - Posy(J) - 1)
    End If
Next
End Function

Function PosChrUCasFst%(S, Optional PosFm% = 1)
Dim J%: For J = PosFm To Len(S)
    If IsUCas(Mid(S, J, 1)) Then PosChrUCasFst = J: Exit Function
Next
End Function
Function PosChrFstNonUCas%(S, Optional PosFm% = 1)
Dim J%: For J = PosFm To Len(S)
    If Not IsUCas(Mid(S, J, 1)) Then PosChrFstNonUCas = J: Exit Function
Next
End Function

Function PosAtDif&(A$, B$) ' Position that A & B has first dif char
Dim LA&, LB&
LA = Len(A)
LB = Len(B)
Dim O&: For O = 1 To Min(LA, LB)
    If Mid(A, O, 1) <> Mid(B, O, 1) Then PosAtDif = O: Exit Function
Next
If LA > LB Then
    PosAtDif = LB + 1
Else
    PosAtDif = LA + 1
End If
End Function

Function PosSsub%(S, Ssub, Optional C As eCas, Optional PosBeg = 1)
If PosBeg < 1 Then ThwPm CSub, "PosBeg must be >=", "S Ssub C PosBeg", S, Ssub, EnmsCas(C), PosBeg
PosSsub = InStr(PosBeg, S, Ssub, VbCprMth(C))
End Function

Function PosySsub(S, Ssub, Optional C As eCas) As Integer()
Const CSub$ = CMod & "PosySsub"
Dim P%: P = 1
Dim M%, J%, L%
L = Len(Ssub)
Again:
    ThwLoopTooMuch CSub, J
    M = PosSsub(S, Ssub, C, P): If M = 0 Then Exit Function
    P = M + L
    PushI PosySsub, M
    GoTo Again
End Function


Private Sub B_PosSsub()
GoSub T1
GoSub T2
GoSub T3
GoSub T4
Exit Sub
Dim S, Ssub, PosBeg%, C As eCas
T1:
    '    12345678901234
    S = ".aaaa.aaaa.bbb"
    Ssub = "."
    PosBeg = 1
    Ept = 1
    GoTo Tst
T2:
    '    12345678901234
    S = ".aaaa.aaaa.bbb"
    Ssub = "."
    PosBeg = 2
    Ept = 6
    GoTo Tst
T3:
    '    12345678901234
    S = ".aaaa.aaaa.bbb"
    Ssub = "."
    PosBeg = 3
    Ept = 11
    GoTo Tst
T4:
    '    12345678901234
    S = ".aaaa.aaaa.bbb"
    Ssub = "."
    PosBeg = 4
    Ept = 0
    GoTo Tst
Tst:
    Act = PosSsub(S, Ssub, C, PosBeg)
    Ass Ept = Act
    Return
End Sub
