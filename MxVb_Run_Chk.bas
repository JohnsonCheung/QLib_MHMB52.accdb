Attribute VB_Name = "MxVb_Run_Chk"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Run_Chk."

Sub ChkNoEr(Er$(), Fun$)
If Si(Er) = 0 Then Exit Sub
Thw Fun, JnCrLf(Er)
End Sub

Sub ChkIsStr(A, Fun$)
If IsStr(A) Then Exit Sub
Thw Fun, "Given parameter should be str, but now TypeName=" & TypeName(A)
End Sub
Sub ChkEq(A, B, Optional P12$ = "Ept Act")
Const CSub$ = CMod & "ChkEq"
ChkIsEqTyn A, B, P12, WHdrChkEq(P12, "TypeName")
Select Case True
Case IsLines(A), IsLines(B)
                  ChkIsEqLines CStr(A), CStr(B), P12, WHdrChkEq(P12, "Lines"): Exit Sub
Case IsStr(A):    ChkIsEqStr CStr(A), CStr(B), P12, WHdrChkEq(P12, "String"):  Exit Sub
Case IsDi(A):     ChkIsEqDi CvDi(A), CvDi(B), P12, WHdrChkEq(P12, "Dictiontary"):  Exit Sub
Case IsArray(A):  ChkIsEqAy A, B, P12, WHdrChkEq(P12, "Array"): Exit Sub
Case IsObject(A): ChkIsEqObj A, B, P12, WHdrChkEq(P12, "Object"): Exit Sub
Case Else:
    If A <> B Then
        Thw CSub, "A B NE", "P12 A B", P12, A, B
        Exit Sub
    End If
End Select
End Sub
Private Function WHdrChkEq$(P12$, Tyn$)
With BrkTRst(P12)
WHdrChkEq = FmtQQ("(?) & (?) of Tyn(?) are NOT EQUAL", .S1, .S2, Tyn)
End With
End Function

Sub ChkIsEqSi(AyA, AyB, Optional Fun$ = "ChkIsEqSi")
If Si(AyA) <> Si(AyB) Then Raise Fun & ": Two array are dif size"
End Sub

Private Sub B_ChkIsEqAy(): ChkIsEqAy Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4): End Sub
Sub ChkIsEqAy(A, B, Optional N12$ = "Ay1 Ay2", Optional Fun$ = "ChkIsEqAy")
If IsEqAy(A, B) Then Exit Sub
Dim N1$, N2$: AsgTAp N12, N1, N2
Dim LinesA$: LinesA = StrTySizAy(A, N1) & vbCrLf & LinesAy(A)
Dim LinesB$: LinesB = StrTySizAy(B, N2) & vbCrLf & LinesAy(B)
CprLines LinesA, LinesB, N12, "Two Ay are different"
Thw CSub, "Two Ay are different"
End Sub
Function StrTySizAy$(Ay, Nm$): StrTySizAy = "-AyTy " & TypeName(Ay) & " -Si " & Si(Ay): End Function
Sub ChkIsEqTyn(A, B, Optional P12$, Optional Hdr$)
Const CSub$ = CMod & "ChkIsEqTyn"
If TypeName(A) = TypeName(B) Then Exit Sub
Dim N$
With Brk1Spc(P12)
    N = FmtQQ("?/?", .S1, .S2)
    Thw CSub, "TypeName of 2 var are Diff:", "Nm Ty1 Ty2", N, TypeName(A), TypeName(B)
End With
End Sub

Sub ChkIsEqAySi(A, B, Optional P12$, Optional Fun$ = "ChkIsEqAySi")
With BrkTRst(P12)
Dim NN$: NN = FmtQQ("[Array 1 Nm(?) size] [Array 2 Nm(?) size]", .S1, .S2)
If Si(A) <> Si(B) Then Thw Fun, "Size-of-2-array ARE DIFfERENT", NN, Si(A), Si(B)
End With
End Sub

Sub ChkHasFf(A As Drs, FF$, Fun$)
If JnSpc(A.Fny) <> FF Then Thw Fun, "Drs-Ff <> Ff", "Drs-Ff Ff", JnSpc(A.Fny), FF
End Sub

Sub ChkIsAySrt(Ay, Fun$)
If IsAySrt(Ay) Then Thw Fun, "Array should be sorted", "Ay-Ty Ay", TypeName(Ay), Ay
End Sub

Sub ChkIsVSomg(A, VarNm$, Fun$)
If IsSomething(A) Then Exit Sub
Thw Fun, FmtQQ("Given[?] is nothing", VarNm)
End Sub

Sub ChkIsPrim(Prim, Optional Fun$ = CMod & "ChkIsPrimy")
If IsPrimy(Prim) Then Exit Sub
Thw Fun, "Given parameter should be prim-array", "Tyn", TypeName(Prim)
End Sub

Function IsPrimy(Ay) As Boolean
If Not IsArray(Ay) Then Exit Function
IsPrimy = IsTynPrim(RmvLas2(TypeName(Ay)))
End Function

Sub ChkIsAy(A, Optional AyNm$ = "Ay", Optional Fun$ = "ChkIsAy")
If IsArray(A) Then Exit Sub
Thw Fun, "Given parameter should be array", "AyNm Tyn", AyNm, TypeName(A)
End Sub
Sub ChkIsEqObj(A, B, Fun$, Optional Msg$ = "Two given object cannot be same")
If Not IsEqObj(A, B) Then Thw Fun, Msg
End Sub

Sub ChkIsIxyIRg(Ixy, Fun$)
Dim O$()
    Dim I, J&: For Each I In Itr(Ixy)
        If I < 0 Then
            PushI O, J & ": " & I
            J = J + 1
        End If
    Next
If Si(O) > 0 Then
    Thw Fun, "In [Ixy], there are [negative-element (Ix Ele)]", "Ixy Neg-Ele", Ixy, O
End If
End Sub
