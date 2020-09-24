Attribute VB_Name = "MxVb_Str"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str."
Function BktIf$(B As Boolean)
If B Then BktIf = "()"
End Function
Function StrTrue$(B As Boolean, S)
If B Then StrTrue = S
End Function
Function StrIfFalse$(B As Boolean, S)
If B = False Then StrIfFalse = S
End Function
Function StrDotIf$(S$)
If S = "" Then StrDotIf = "." Else StrDotIf = S
End Function

Function Pad0$(N, NDig):  Pad0 = Format(N, StrDup("0", NDig)): End Function
Function Pad02$(N):      Pad02 = Pad0(N, 2):                   End Function
Function Pad04$(N):      Pad04 = Pad0(N, 4):                   End Function
Function StrDup$(S, N)
Dim O$, J&
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function


Sub EdtStr(S, Ft)
WrtStr S, Ft, OvrWrt:=True
Brw Ft
End Sub

Function IsStrDig(S) As Boolean
Dim J&
For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsStrDig = True
End Function

Function LeftOrAll(S, L%)
If L <= 0 Then
    LeftOrAll = S
Else
    LeftOrAll = Left(S, L)
End If
End Function

Function PosyQuoDbl(S) As Integer(): PosyQuoDbl = PosySsub(S, vbQuoDbl): End Function

Function IsLnInd(Ln) As Boolean: IsLnInd = LTrim(ChrFst(Ln)) = "": End Function ' ret True if after Trim fst chr is a ""
Function AliR5$(V): AliR5 = AliMax(V, 5, IsAliR:=True): End Function
Function AliR7$(V): AliR7 = AliMax(V, 7, IsAliR:=True): End Function
Function AliR4$(V): AliR4 = AliMax(V, 4, IsAliR:=True): End Function
Function AliR3$(V): AliR3 = AliMax(V, 3, IsAliR:=True): End Function
Function AliR2$(V): AliR2 = AliMax(V, 2, IsAliR:=True): End Function

Function AliMax$(V, W%, Optional IsAliR As Boolean)
AliMax = Ali(V, W, IsAliR)
If Len(AliMax) > W Then
    If W > 4 Then
        AliMax = Left(AliMax, W - 4) & "..."
    Else
        AliMax = RmvLas(AliMax) & "*"
    End If
End If
End Function
Function Ali$(V, W%, IsAliR As Boolean)
If IsAliR Then
    Ali = AliR(V, W)
Else
    Ali = AliL(V, W)
End If
End Function
Function AliL$(S, W%)
If W <= 0 Then Exit Function
Dim L%: L = Len(S)
If L >= W Then
    AliL = S
Else
    AliL = S & Space(W - Len(S))
End If
End Function
Function AliR$(S, W)
Dim L%: L = Len(S)
If W > L Then
    AliR = Space(W - L) & S
Else
    AliR = S
End If
End Function


Function TabN$(N%): TabN = Space(4 * N): End Function
