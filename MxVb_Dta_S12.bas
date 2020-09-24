Attribute VB_Name = "MxVb_Dta_S12"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_S12."
Type S123: S1 As String: S2 As String: S3 As String: End Type 'Deriving(Ctor)
Type S12: S1 As String: S2 As String: End Type 'Deriving(Ay Ctor Opt)
Type OptS12: Som As Boolean: S12 As S12: End Type
Function S123(S1, S2, S3) As S123
With S123
    .S1 = S1
    .S2 = S2
    .S3 = S3
End With
End Function
Function S12yRmvBlnkS2(S() As S12) As S12()
Dim J%: For J = 0 To UbS12(S)
    If Trim(S(J).S2) <> "" Then PushS12 S12yRmvBlnkS2, S(J)
Next
End Function

Sub AsgS12(S As S12, O1, O2): O1 = S.S1: O2 = S.S2: End Sub

Private Sub B_S12yDi()
Dim D As New Dictionary, Act() As S12, Ept() As S12
D.Add "A", "BB"
D.Add "B", "CCC"
Act = S12yDi(D)
Stop
End Sub

Function DiS12yS1Unq(A() As S12) As Dictionary
Set DiS12yS1Unq = New Dictionary
Dim J&: For J = 0 To UbS12(A)
    PushDiS12 DiS12yS1Unq, A(J)
Next
End Function
Function DiS12yS1Dup(A() As S12, Optional Sep$ = vbCrLf) As Dictionary ' ret a di from @S with S1 can be dup,
' so that the S2 of the same S1 will be concat up to a string as the returned dic-Val by using @Sep
Dim O As New Dictionary
Dim J&: For J = 0 To UbS12(A)
    With A(J)
        If O.Exists(.S1) Then
            O(.S1) = O(.S1) & Sep & .S2
        Else
            PushDiS12 O, A(J)
        End If
    End With
Next
End Function
Function SqS12y(S() As S12, Optional H12$ = "S1 S2") As Variant()
Dim N&: N = SiS12(S)
If N = 0 Then Exit Function
Dim O(), I, R&, J&
ReDim O(1 To N, 1 To 2)
Dim N1$, N2$
AsgTAp H12, N1, N2
R = 2
O(1, 1) = N1
O(1, 2) = N2
For J = 0 To N - 1
    With S(J)
        O(R, 1) = .S1
        O(R, 2) = .S2
        R = R + 1
    End With
Next
SqS12y = O
End Function
Function S12ySyBrk1(Sy$(), StrBrk$) As S12()
Dim S: For Each S In Itr(Sy)
    PushS12 S12ySyBrk1, Brk1(S, StrBrk)
Next
End Function
Function S12ySyBrk2(Sy$(), StrBrk$) As S12()
Dim S: For Each S In Itr(Sy)
    PushS12 S12ySyBrk2, Brk2(S, StrBrk)
Next
End Function

Function S12yAy12(A, B, Optional NoTrim As Boolean) As S12()
ChkIsEqAySi A, B, , CSub
Dim U&, O() As S12
U = UB(A)
ReDim O(U)
Dim J&: For J = 0 To U
    O(J) = S12Nw(A(J), B(J), NoTrim)
Next
S12yAy12 = O
End Function
Function S12yDi(D As Dictionary) As S12()
Dim K: For Each K In D.Keys
    PushS12 S12yDi, S12(K, D(K))
Next
End Function

Function S1y(A() As S12) As String()
Dim J&: For J = 0 To UbS12(A)
   PushI S1y, A(J).S1
Next
End Function
Function S2y(A() As S12) As String()
Dim J&: For J = 0 To UbS12(A)
   Push S2y, A(J).S2
Next
End Function

Function S12yDif(A() As S12, B() As S12) As S12()
'Ret : Subset of @A.  Those itm in @A also in @B will be exl.
Dim J&: For J = 0 To UbS12(A)
    If Not HasS12(B, A(J)) Then
        PushS12 S12yDif, A(J)
    End If
Next
End Function
Function IsEqS12(S As S12, B As S12) As Boolean
With S
    If .S1 <> B.S1 Then Exit Function
    If .S2 <> B.S2 Then Exit Function
End With
IsEqS12 = True
End Function
Function HasS12(A() As S12, B As S12) As Boolean
Dim J&: For J = 0 To SiS12(A)
    If IsEqS12(A(J), B) Then HasS12 = True: Exit Function
Next
End Function
Function HasS1(A() As S12, S1$) As Boolean
Dim J&: For J = 0 To SiS12(A)
    If A(J).S1 = S1 Then HasS1 = True: Exit Function
Next
End Function
Function AddS2Sfx(A() As S12, S2Sfx$) As S12()
Dim O() As S12: O = A
Dim J&: For J = 0 To UbS12(A)
    O(J).S2 = O(J).S2 & S2Sfx
Next
AddS2Sfx = O
End Function
Function S12Nw(S1, S2, Optional NoTrim As Boolean) As S12
S12Nw = S12(S1, S2)
If Not NoTrim Then S12Nw = S12Trim(S12Nw)
End Function
Function S12Trim(S As S12) As S12: S12Trim = S12(Trim(S.S1), Trim(S.S2)): End Function
Function MapS1(A() As S12, Dic As Dictionary) As S12()
Const CSub$ = CMod & "MapS1"
Dim J&: For J = 0 To UbS12(A)
    Dim M As S12: M = A(J)
    If Not Dic.Exists(M.S1) Then
        Thw CSub, "Som S1 in [S12y] not found in [Dic]", "S1-not-found S12y Dic", M.S1, FmtS12y(A), FmtDi(Dic)
    End If
    M.S1 = Dic(M.S1)
    PushS12 MapS1, M
Next
End Function
Sub WrtS12y(S() As S12, Ft$, Optional OvrWrt As Boolean):          WrtStr S12lS12y(S), Ft, OvrWrt: End Sub
Function S12yFt(Ft) As S12():                             S12yFt = S12yS12l(LinesFt(Ft)):          End Function
Function SwapS12y(A() As S12) As S12()
Dim O() As S12: O = A
Dim J&: For J = 1 To UbS12(A)
    O(J) = S12Swap(A(J))
Next
SwapS12y = O
End Function
Sub PushSPr(O() As S12, S1$, S2$, Optional NoTrim As Boolean): PushS12 O, S12Nw(S1, S2, NoTrim): End Sub
Function S12yAddS1Pfx(A() As S12, S1Pfx$) As S12()
Dim J&: For J = 0 To UbS12(A)
    Dim M As S12: M = A(J)
    M.S1 = S1Pfx & M.S1
    PushS12 S12yAddS1Pfx, M
Next
End Function
Sub PushS12NBS1(O() As S12, M As S12)
If M.S1 <> "" Then PushS12 O, M
End Sub
Sub PushS12Opt(O() As S12, M As OptS12)
If M.Som Then PushS12 O, M.S12
End Sub
Sub PushDiS12(O As Dictionary, M As S12, Optional Sep$ = vbCrLf)
With M
    If O.Exists(.S1) Then
        O(.S1) = O(.S1) & " " & O(.S2)
    Else
        O.Add .S1, .S2
    End If
End With
End Sub

Function S12yDrs(D As Drs, Optional P12$) As S12()
'Fm D  : ..@P12.. ! A drs with col-@P12.  At least has 2 col
'Fm P12 :         ! if isBlnk, use fst 2 col
'Ret   :         ! fst col will be S1 and snd col will be S2 join with vbCrLf
Dim S1$(), S2() ' S2 is ay of sy
Dim I1%, I2%
    If P12 = "" Then I1 = 0: I2 = 1 Else AsgCxapDrs D, P12, I1, I2
Dim Dr: For Each Dr In Itr(D.Dy)
    Dim A$, B$: A = Dr(I1): B = Dr(I2)
    Dim R&: R = IxiEle(S1, A)
    If R = -1 Then
        PushI S1, A
        PushI S2, Sy(B)
    Else
        PushI S2(R), B
    End If
Next
Dim J&: For J = 0 To UB(S1)
    PushS12 S12yDrs, S12(S1(J), JnCrLf(S2(J)))
Next
End Function
Function JnS12$(S As S12, Optional Brk$): JnS12 = S.S1 & Brk & S.S2: End Function
Function STup2S12(S As S12) As String()
Dim O$(): ReDim O(1)
O(0) = S.S1
O(1) = S.S2
STup2S12 = O
End Function
Function S12Swap(S As S12) As S12: S12Swap = S12(S.S2, S.S1): End Function
Function FstS2WhS1$(S() As S12, S1$)
Dim J%: For J = 0 To UbS12(S)
    With S(J)
        If .S1 = S1 Then FstS2WhS1 = .S2: Exit Function
    End With
Next
End Function
Function S12(S1, S2) As S12
With S12
    .S1 = S1
    .S2 = S2
End With
End Function
Function S12yAdd(S As S12, B As S12) As S12(): PushS12 S12yAdd, S: PushS12 S12yAdd, B: End Function
Sub PushS12y(O() As S12, A() As S12): Dim J&: For J = 0 To UbS12(A): PushS12 O, A(J): Next: End Sub
Sub PushS12(O() As S12, M As S12): Dim N&: N = SiS12(O): ReDim Preserve O(N): O(N) = M: End Sub
Sub PushS12Som(O() As S12, M As OptS12)
With M
    If .Som Then PushS12 O, .S12
End With
End Sub
Function SiS12&(A() As S12): On Error Resume Next: SiS12 = UBound(A) + 1: End Function
Function UbS12&(A() As S12): UbS12 = SiS12(A) - 1: End Function
Function OptS12(Som, S As S12) As OptS12: With OptS12: .Som = Som: .S12 = S: End With: End Function
Function SomS12(S As S12) As OptS12: SomS12.Som = True: SomS12.S12 = S: End Function

Function S12yLy(Ly$(), Sep$, Optional NoTrim As Boolean) As S12()
Dim O() As S12
Dim U&: U = UB(Ly)
ReDim O(U)
Dim J&: For J = 0 To U
    O(J) = Brk1(Ly(J), Sep, NoTrim)
Next
S12yLy = O
End Function
