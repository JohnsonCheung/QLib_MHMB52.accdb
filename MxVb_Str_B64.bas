Attribute VB_Name = "MxVb_Str_B64"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_B64."
Const Str64$ = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Const NChrB64% = 128
Const NChrStr% = 128 / 4 * 3
Private X_1Bin4y$() ' 16 ele each is Len-4 from 0000 to 1111
Private X_3Bin6y$() ' 64 ele each is Len-6 from 000000 to 111111
Private Sub B_B64()
'GoSub T1
'GoSub T2
GoSub ZZ
Exit Sub
Dim Str$
T2:
    Str = "Ma"
    Ept = "TWE="
    GoTo Tst
T1:
    Str = "M"
    Ept = "TQ=="
    GoTo Tst
Tst:
    Act = B64(Str)
    C
    Return
ZZ: Vc B64(SrclPC)
    Return
End Sub
Function B64$(Str$)
If Si(X_1Bin4y) = 0 Then X_1Bin4y = WBin4y
Dim L&: L = Len(Str)
Dim O$()
    Dim J&: For J = 1 To L - (L Mod 3) Step 3
        PushI O, WB644(Mid(Str, J, 3))
    Next
    Select Case L Mod 3
    Case 1: PushI O, WB644Chr1(ChrLas(Str))
    Case 2: PushI O, WB644Chr2(Right2(Str))
    End Select
B64 = Jn(O)
End Function
Private Sub B_WBin4y(): VcAy WBin4y: End Sub
Private Function WBin4y() As String()
Const S$ = "0000 0001 0010 0011 0100 0101 0110 0111" & _
          " 1000 1001 1010 1011 1100 1101 1110 1111"
WBin4y = SplitSpc(S)
End Function
Private Sub B_WB644()
X_1Bin4y = WBin4y
GoSub T1
GoSub T2
Exit Sub
Dim Left3 As String * 3
T1:
    Left3 = "Man"
    Ept = "TWFu"
    GoTo Tst
T2:
    Left3 = "abc"
    Ept = "YWJj"
    GoTo Tst
Tst:
    Act = WB644(Left3)
    If Act <> Ept Then Stop
    C
    Return
End Sub
Private Function WB644$(Left3$)
Dim Bin24$
    Dim C1$, C2$, C3$
    C1 = ChrFst(Left3)
    C2 = Mid(Left3, 2, 1)
    C3 = ChrLas(Left3)
    Bin24 = X_1Bin8(C1) & X_1Bin8(C2) & X_1Bin8(C3)
Dim B64y$(3)
Dim J As Byte: For J = 0 To 3
    Dim Bin6$: Bin6 = Mid(Bin24, J * 6 + 1, 6)
    B64y(J) = X_1B64(Bin6)
Next
WB644 = Jn(B64y)
End Function
Private Function WB644Chr1(Chr1$)
Dim Bin12$: Bin12 = X_1Bin8(Chr1) & "0000"
Dim B64y$(1)
Dim J As Byte: For J = 0 To 1
    Dim Bin6$: Bin6 = Mid(Bin12, J * 6 + 1, 6)
    B64y(J) = X_1B64(Bin6)
Next
WB644Chr1 = Jn(B64y) & "=="
End Function
Private Function WB644Chr2$(ChrSnd$)
Dim C1$, C2$
    C1 = ChrFst(ChrSnd)
    C2 = ChrLas(ChrSnd)
Dim Bin18$: Bin18 = X_1Bin8(C1) & X_1Bin8(C2) & "00"
Dim B64y$(2)
Dim J As Byte: For J = 0 To 2
    Dim Bin6$: Bin6 = Mid(Bin18, J * 6 + 1, 6)
    B64y(J) = X_1B64(Bin6)
Next
WB644Chr2 = Jn(B64y) & "="
End Function

Private Function X_1B64$(Bin6$)
Dim Pos64 As Byte
    Dim Ix4A As Byte: Ix4A = Ix4Bin2(Left(Bin6, 2))
    Dim Ix4B As Byte: Ix4B = Ix4Bin2(Mid(Bin6, 3, 2))
    Dim Ix4C As Byte: Ix4C = Ix4Bin2(Right(Bin6, 2))
    Pos64 = Ix4A * 16 + Ix4B * 4 + Ix4C + 1
X_1B64 = Mid(Str64, Pos64, 1)
End Function

Private Sub B_X_1Bin8()
X_1Bin4y = WBin4y
Dim O$()
Dim J%: For J = 0 To 255
    PushI O, Hex2(Chr(J)) & " " & X_1Bin8(Chr(J))
Next
VcAy O
End Sub
Private Function X_1Bin8$(C$)
Dim A%: A = Asc(C)
X_1Bin8 = X_1Bin4y(A \ 16) & X_1Bin4y(A Mod 16)
End Function
Private Sub B_StrB64()
GoSub T1
GoSub T2
GoSub T3
GoSub T4
GoSub Z1
Exit Sub
Dim B64_$
Dim Str$, AA_B64$, AA_Str$
T1:
    Ept = "Man"
    B64_ = "TWFu"
    GoTo Tst
T2:
    B64_ = "YWJj"
    Ept = "abc"
    GoTo Tst
T3:
    B64_ = "TWE="
    Ept = "Ma"
    GoTo Tst
T4:
    Ept = "M"
    B64_ = "TQ=="
    GoTo Tst
ZZ: Vc B64(SrclPC)
    Return

Tst:
    Act = StrB64(B64_)
    C
    Return
Z1:
    Str = "Man":   GoSub TstZZ
    Str = "Ma":    GoSub TstZZ
    Str = "M":     GoSub TstZZ
    Str = SrclPC:  GoSub TstZZ
Z2:
    Str = SrclMC
    GoTo TstZZ
TstZZ:
    AA_B64 = B64(Str)
    AA_Str = StrB64(AA_B64)
    If Str <> AA_Str Then Stop
    Return
End Sub
Function StrB64$(B64$)
Const CSub$ = CMod & "StrB64"
If Si(X_3Bin6y) = 0 Then X_3Bin6y = W3Bin6y
Dim L&: L = Len(B64)
If L Mod 4 <> 0 Then Thw CSub, "Len of B64 mod 4 <> 0", "[Mod 4 of Len(@B64)] Len(@B64)", L Mod 4, L
Dim O$()
Dim JLas&: JLas = ((L \ 4) - 1) * 4 + 1
Dim J&: For J = 1 To L Step 4
    Dim B644$: B644 = Mid(B64, J, 4)
    If J = JLas Then
        PushI O, W3Str1or2or3(B644)
    Else
        PushI O, W3Str3(B644)
    End If
Next
StrB64 = Jn(O)
End Function
Private Sub B_W3Bin6y(): VcAy W3Bin6y: End Sub
Private Function W3Bin6y() As String()
Const S$ = _
 "000000 000001 000010 000011 000100 000101 000110 000111" & _
" 001000 001001 001010 001011 001100 001101 001110 001111" & _
" 010000 010001 010010 010011 010100 010101 010110 010111" & _
" 011000 011001 011010 011011 011100 011101 011110 011111" & _
" 100000 100001 100010 100011 100100 100101 100110 100111" & _
" 101000 101001 101010 101011 101100 101101 101110 101111" & _
" 110000 110001 110010 110011 110100 110101 110110 110111" & _
" 111000 111001 111010 111011 111100 111101 111110 111111"
W3Bin6y = SplitSpc(S)
End Function
Private Function W3Str3$(B644$)
Dim A$, B$, C$, D$
    A = ChrFst(B644)
    B = Mid(B644, 2, 1)
    C = Mid(B644, 3, 1)
    D = ChrLas(B644)
Dim Bin24$: Bin24 = X_3Bin6(A) & X_3Bin6(B) & X_3Bin6(C) & X_3Bin6(D)
Dim Bin8A$: Bin8A = Left(Bin24, 8)
Dim Bin8B$: Bin8B = Mid(Bin24, 9, 8)
Dim Bin8C$: Bin8C = Right(Bin24, 8)
W3Str3 = X_3Chr(Bin8A) & X_3Chr(Bin8B) & X_3Chr(Bin8C)
End Function
Private Function W3Str1or2or3$(B644Las$)
Select Case True
Case Right2(B644Las) = "==": W3Str1or2or3 = W3Str1(Left2(B644Las))
Case ChrLas(B644Las) = "=":   W3Str1or2or3 = W3Str2(Left3(B644Las))
Case Else:                    W3Str1or2or3 = W3Str3(B644Las)
End Select
End Function
Private Function W3Str1$(B642$)
Dim A$, B$
    A = ChrFst(B642)
    B = ChrLas(B642)
Dim Bin12$: Bin12 = X_3Bin6(A) & X_3Bin6(B)
Dim Bin8$: Bin8 = Left(Bin12, 8)
W3Str1 = X_3Chr(Bin8)
End Function
Private Function W3Str2$(B643$)
Dim A$, B$, C$
    A = ChrFst(B643)
    B = Mid(B643, 2, 1)
    C = ChrLas(B643)
Dim Bin18$: Bin18 = X_3Bin6(A) & X_3Bin6(B) & X_3Bin6(C)
Dim Bin8A$: Bin8A = Left(Bin18, 8)
Dim Bin8B$: Bin8B = Mid(Bin18, 9, 8)
W3Str2 = X_3Chr(Bin8A) & X_3Chr(Bin8B)
End Function

Private Function X_3Chr$(Bin8$)
Dim A As Byte
    Dim Ix4A As Byte: Ix4A = Ix4Bin2(Left(Bin8, 2))
    Dim Ix4B As Byte: Ix4B = Ix4Bin2(Mid(Bin8, 3, 2))
    Dim Ix4C As Byte: Ix4C = Ix4Bin2(Mid(Bin8, 5, 2))
    Dim Ix4D As Byte: Ix4D = Ix4Bin2(Right(Bin8, 2))
    A = Ix4A * 64 + Ix4B * 16 + Ix4C * 4 + Ix4D
X_3Chr = Chr(A)
End Function
Private Function X_3Bin6$(B641$)
Const CSub$ = CMod & "X_3Bin6"
Dim Ix64 As Byte
    Ix64 = InStr(1, Str64, B641, vbBinaryCompare): If Ix64 = 0 Then Thw CSub, "B641 has invalid value", "Asc(B641) B641", Asc(B641), B641
    Ix64 = Ix64 - 1
Dim Ix8A As Byte, Ix8B As Byte
    Ix8A = Ix64 \ 8
    Ix8B = Ix64 Mod 8
X_3Bin6 = X_3Bin6y(Ix64)
End Function
