Attribute VB_Name = "MxVb_Str_Nm"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Nm."
Private Sub B_ShfNm()
GoSub T1
Exit Sub
Dim Ln$, EptLn$
T1:
    Ln = "aa  bb"
    Ept = "aa"
    EptLn = "bb"
    GoTo Tst
Tst:
    Act = ShfNm(Ln)
    Debug.Assert Act = Ept
    Debug.Assert Ln = EptLn
    Return
End Sub
Function IsNm(S) As Boolean
If S = "" Then Exit Function
If Not IsLetter(ChrFst(S)) Then Exit Function
Dim L&: L = Len(S)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsChrNm(Mid(S, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsChrNm(C$) As Boolean
IsChrNm = True
If IsLetter(C) Then Exit Function
If C = "_" Then Exit Function
If IsDig(C) Then Exit Function
IsChrNm = False
End Function

Function IsChrDotNm(A$) As Boolean
If IsChrNm(A) Then IsChrDotNm = True: Exit Function
IsChrDotNm = A = "."
End Function

Function TakNmDot$(S, Optional PosBeg% = 1)
Dim L%: L = Len(S): If L = 0 Then Exit Function
If Not IsLetter(ChrFst(S)) Then Exit Function
Dim J%: For J = 2 To L
    If Not IsChrDotNm(Mid(S, J, 1)) Then
        TakNmDot = Left(S, J - 1)
        Exit Function
    End If
Next
TakNmDot = S
End Function

Function AftNm$(S): AftNm = Mid(S, Len(TakNm(S)) + 1): End Function
Function TakNm$(S, Optional PosBeg% = 1)
If Mid(S, PosBeg) = "" Then Exit Function
If Not IsLetter(ChrAt(S, PosBeg)) Then Exit Function
Dim J%: For J = PosBeg + 1 To Len(S)
    If Not IsChrNm(Mid(S, J, 1)) Then
        TakNm = Left(S, J - 1)
        Exit Function
    End If
Next
TakNm = S
End Function
Sub ChkNy(Ny$(), Fun$)
Dim N: For Each N In Itr(Ny)
    If Not IsNm(N) Then Thw Fun, "Ele of Sy is not nm", "Not-nm-Ele Sy", N, Sy
Next
End Sub
Function ShfNmMayBktSq$(OLn$)
Dim O$
If ChrFst(OLn) = "[" Then
    Dim P%: P = InStr(OLn, "]"): If P = 0 Then Exit Function
    ShfNmMayBktSq = RmvFst(Left(OLn, P - 1))
    OLn = Mid(OLn, P + 1)
Else
    ShfNmMayBktSq = ShfNm(OLn)
End If
End Function

Function ShfNm$(OLn$)
Dim O$: O = TakNm(OLn): If O = "" Then Exit Function
ShfNm = O
OLn = LTrim(RmvPfx(OLn, O))
End Function

Function ShfNmDot$(OLn$)
Dim O$: O = TakNmDot(OLn): If O = "" Then Exit Function
ShfNmDot = O
End Function

Function NnWhLikk$(NN$, WhLikk$): NnWhLikk = JnSpc(NyWhLiky(SplitSpc(NN), SplitSpc(WhLikk))): End Function
Function NyWhLiky(Ny$(), Liky$()) As String() 'Return SubSet of @Ny for each ele is @WhLikNy
Dim N: For Each N In Itr(Ny)
    If HitLiky(N, Liky) Then PushI NyWhLiky, N
Next
End Function


Function IsNmInst(Nm) As Boolean
If ChrFst(Nm) <> "N" Then Exit Function      'ChrFst = N
If Len(Nm) <> 16 Then Exit Function          'Len    =16
If Not IsYYYYMMDD(Mid(Nm, 2, 8)) Then Exit Function 'NYYYYMMDD_HHMMDD
If Mid(Nm, 10, 1) <> "_" Then Exit Function
If Not IsHHMMDD(Right(Nm, 6)) Then Exit Function
IsNmInst = True
End Function
