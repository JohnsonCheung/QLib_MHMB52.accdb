Attribute VB_Name = "MxVb_Str_Wrd_Wrdln"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_WrpWrd."

Private Sub B_WrdlLines():                                                 Vc WrdlLines(SrclPC, 80):        End Sub
Function WrdlLines$(Lines, Optional Wdt% = 80):                WrdlLines = JnCrLf(WrdlnyLines(Lines, Wdt)): End Function
Function WrdlnyLines(Lines, Optional Wdt% = 80) As String(): WrdlnyLines = WrdlnyLy(SplitCrLf(Lines), Wdt): End Function
Function WrdlnyLy(Ly$(), Optional Wdt% = 80) As String()
Dim Ln: For Each Ln In Itr(Ly)
    PushIAy WrdlnyLy, WrdlnyLn(Ln, Wdt)
Next
End Function

Private Sub B_WrdlnyLn()
GoSub T0
GoSub T1
GoSub T2
Exit Sub
Dim W%, L$
T0:
    W = 80
    L = "slkdfj sldkjf sdklj dklfjdf"
    Ept = Sy(L)
    GoTo Tst
T1:
    W = 80
    L = "AddColMthl DoCachedMthcP DoCachedMthcP DrsTMth DrsTMthc DrsTMthcM DrsTMthcP DrsTMthcP__Tst DrsTMthcFxa DrsTMthcM DrsTMthcP" & _
        " DrsTMthcPjf DrsTMthcPjfy DrsTMthcV DrsTMthM DrsTMthP DrsTMthM DrsTMthP PFunDrsP MthDr MthlnDr WsoMthcP B_DrsTMthcP"
    Ept = Sy("AddColMthl DoCachedMthcP DoCachedMthcP DrsTMth DrsTMthc DrsTMthcM DrsTMthcP", _
             ". DrsTMthcP__Tst DrsTMthcFxa DrsTMthcM DrsTMthcP DrsTMthcPjf DrsTMthcPjfy", _
             ". DrsTMthcV DrsTMthM DrsTMthP DrsTMthM DrsTMthP PFunDrsP MthDr MthlnDr WsoMthcP", _
             ". B_DrsTMthcP")
    GoTo Tst
T2:
    W = 40
    L = "AddColMthl DoCachedMthcP DoCachedMthcP DrsTMth DrsTMthc DrsTMthcM DrsTMthcP DrsTMthcP__Tst DrsTMthcFxa DrsTMthcM DrsTMthcP" & _
        " DrsTMthcPjf DrsTMthcPjfy DrsTMthcV DrsTMthM DrsTMthP DrsTMthM DrsTMthP PFunDrsP MthDr MthlnDr WsoMthcP B_DrsTMthcP"
    Ept = Sy("AddColMthl DoCachedMthcP DoCachedMthcP", _
            ". DrsTMth DrsTMthc DrsTMthcM DrsTMthcP", _
            ". DrsTMthcP__Tst DrsTMthcFxa DrsTMthcM", _
            ". DrsTMthcP DrsTMthcPjf DrsTMthcPjfy", _
            ". DrsTMthcV DrsTMthM DrsTMthP DrsTMthM", _
            ". DrsTMthP PFunDrsP MthDr MthlnDr", _
            ". WsoMthcP B_DrsTMthcP")
    GoTo Tst
Tst:
    Act = WrdlnyLn(L, W)
'    Dmp Act
    C
    Return
End Sub
Private Function WrdlnyLn(L, Optional W% = 80) As String()
Dim S$: S = RTrim(L)
Dim O$()
PushNB O, ShfWrdln(S, W)
Dim W1%: W1 = W - 2
Dim J%: While S <> ""
    ThwLoopTooMuch CSub, J
    Dim A$: A = ShfWrdln(S, W1)
    If A = "" Then
        If S = "" Then Exit Function
        Thw CSub, "After !W0_ShfLn, Shifted Str is blnk but the remaining is NB.", "[@Orig Ln] @W [Remainging str after a shifted string is blank]", L, W, S
    End If
    PushI O, ". " & A
Wend
WrdlnyLn = O
End Function
Private Sub B_ShfWrdln()
GoSub T1
GoSub T2
Exit Sub
Dim OLn$, W%, OLnEpt$
T1:
    OLn = "1234 678 "
    W = 7
    Ept = "1234"
    OLnEpt = "678 "
    GoTo Tst
T2:
    OLn = "1234 678"
    W = 3
    Ept = "123"
    OLnEpt = "4 678"
    GoTo Tst
Tst:
    Act = ShfWrdln(OLn, W%)
    C
    Ass OLnEpt = OLn
    Return
End Sub
Private Function ShfWrdln$(OLn$, W%) ' EleShf a Ln from @Ln, at most the lenght is @W
If Len(OLn) < W Then
    ShfWrdln = OLn
    OLn = ""
    Exit Function
End If
Dim N%: N = NLeftWrdln(OLn, W)
ShfWrdln = RTrim(Left(OLn, N))
OLn = LTrim(Mid(OLn, N + 1))
End Function

Private Function NLeftWrdln%(Ln$, W%)
If IsPosWrdCut(Ln, W) Then
    Dim P&: P = InStrRev(Left(Ln, W), " ")
    NLeftWrdln = IIf(P = 0, W, P)
Else
    NLeftWrdln = W
End If
End Function
Private Function IsPosWrdCut(Ln$, Pos%) As Boolean
Select Case True
Case Mid(Ln, Pos, 1) = " ", Mid(Ln, Pos + 1, 1) = " "
Case Else: IsPosWrdCut = True
End Select
End Function

