Attribute VB_Name = "MxVb_Msg_Box"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Msg_Box."

Function BoxLines(Lines, Optional C$ = "*") As String(): BoxLines = BoxLy(SplitCrLf(Lines), C): End Function
Function BoxAy(Ay, Optional C$ = "*") As String():          BoxAy = BoxLy(SyAy(Ay), C):         End Function
Function BoxLy(Ly$(), Optional C$ = "*") As String()
If Si(Ly) = 0 Then Exit Function
Dim W%, L$, I
W = AyWdt(Ly)
L = StrDup(C, W + 6)
PushI BoxLy, L
Dim Q1$: Q1 = C & C & " "
Dim Q2$: Q2 = " " & C & C
For Each I In Ly
    PushI BoxLy, Q1 & AliL(I, W) & Q2
Next
PushI BoxLy, L
End Function

Function BoxStr(S$, Optional C$ = "*") As String()
If Trim(S) = "" Then Exit Function
Dim H$: H = StrDup(C, Len(S) + 6)
PushI BoxStr, H
PushI BoxStr, C & C & " " & S & " " & C & C
PushI BoxStr, H
End Function

Function Box(V, Optional C$ = "*") As String()
If V = "" Then Exit Function
If IsStr(V) Then
    If V = "" Then
        Exit Function
    End If
End If
Select Case True
Case IsLines(V): Box = BoxLines(V, C)
Case IsStr(V):   Box = BoxStr(CStr(V), C)
Case IsSy(V):    Box = BoxLy(CvSy(Sy), C)
Case IsArray(V): Box = BoxAy(V)
Case Else:       Box = BoxStr(CStr(V), C)
End Select
End Function

Function Boxl$(V, Optional C$ = "*") 'vbCrLf is always at end
If V = "" Then Exit Function
Boxl = JnCrLf(Box(V, C)) & vbCrLf
End Function

Function BoxFny(Fny$()) As String()
If Si(Fny) = 0 Then Exit Function
Const S$ = " | ", Q$ = "| * |"
Const LS$ = "-|-", LQ$ = "|-*-|"
Dim L$, H$, Ay$(), J%
    ReDim Ay(UB(Fny))
    For J = 0 To UB(Fny)
        Ay(J) = StrDup("-", Len(Fny(J)))
    Next
L = Quo(Jn(Fny, S), Q)
H = Quo(Jn(Ay, LS), LQ)
BoxFny = Sy(H, L, H)
End Function

Function BoxFxw(Fx$, W$, Tit$) As String()
Dim B$()
If Tit <> "" Then B = Box(Tit)
BoxFxw = SyAdd(B, MsgFxw(Fx, W))
End Function
