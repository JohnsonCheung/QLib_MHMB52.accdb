Attribute VB_Name = "MxVb_Str_Cpr_Fc_Mul"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Cpr_Fc_Mul."

Private Sub B_LsyFc()
Dim A$(), B$()
A = SplitVBar("ksldfsdf|lksdfjlsdkf|sldfklfs2")
B = SplitVBar("ksldfsdf|lksdfjlsdkf|sldfklfs")
Brw LsyFc(A, B)
End Sub
Function LsyS12yFc(S() As S12) As String():     LsyS12yFc = LsyFc(S1y(S), S2y(S)):                    End Function
Function LyLyy12Fc(Lyy1(), Lyy2()) As String(): LyLyy12Fc = LyLsy(LsyFc(LsyLyy(Lyy1), LsyLyy(Lyy2))): End Function

Function LsyFc(Lsy1$(), Lsy2$()) As String()
Dim Ixy&(): Ixy = WIxy(Lsy1, Lsy2) ' Ixy of diff
If Si(Ixy) = 0 Then LsyFc = Sy(FmtQQ("All [?] pair of lines are same", Si(Lsy1))): Exit Function
Dim P$: P = PthTmpInst("FileCompareLsy(Fc)")
WWrtAB P, Ixy, Lsy1, Lsy2
WWrtCdl P, Ixy
LsyFc = WLsy(P, Ixy, Lsy1, Lsy2)
End Function
Private Function WIxy(A$(), B$()) As Long() ' Ixy of dif
Dim J%: For J = 0 To UB(A)
    If A(J) <> B(J) Then
        PushI WIxy, J
    End If
Next
End Function
Private Function WWrtAB(Pth$, Ixy&(), A$(), B$())
Dim F$: F = String(Len(CStr(UB(A))), "0")
Dim J%: For J = 0 To UB(Ixy)
    Dim Ix&: Ix = Ixy(J)
    Dim Pfx$: Pfx = Pth & Format(Ix, F)
    WrtStr A(Ix), Pfx & ".A.txt"
    WrtStr B(Ix), Pfx & ".B.txt"
Next
End Function
Private Sub WWrtCdl(Pth$, Ixy&())
Dim Cdl$: Cdl = LinesApLn( _
    FmtQQ("Cd ""?""", Pth), _
    WCdl(Ixy), _
    "Time >end.txt")
Shell FcmdCrt("FileCompareLsy", Cdl), vbHide
ChkWaitFfn Pth & "end.txt"
End Sub
Private Function WCdl$(Ixy&())
Dim O$()
    Dim F$: F = String(Len(CStr(EleLas(Ixy))), "0")
    Dim J%: For J = 0 To UB(Ixy)
        Dim P$: P = Format(Ixy(J), F)
        PushI O, RplQ("Fc ?.A.txt ?.B.txt > ?Fc.txt", P)
    Next
WCdl = JnCrLf(O)
End Function
Private Function WLsy(Pth$, Ixy&(), A$(), B$()) As String() ' Read P\000?Fc.Txt"
Dim O$()
Dim F$: F = String(Len(CStr(UB(A))), "0")
Dim J&: For J = 0 To UB(Ixy)
    PushI O, LinesEndTrim(LinesFt(Pth & Format(Ixy(J), F) & "Fc.Txt"))
Next
WLsy = O
End Function
