Attribute VB_Name = "MxVb_Fun"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fun."
Public Const CSub$ = CMod & "?"
Sub Stp(): Stop: End Sub
Sub Swap(OA, OB)
Dim X: X = OA
OA = OB
OB = X
End Sub

Function CUsr$(): CUsr = Environ$("USERNAME"): End Function

Sub Asg(Fm, _
Into)
Select Case True
Case IsObject(Fm): Set Into = Fm
Case Else:             Into = Fm
End Select
End Sub

Sub ShellMaxCdl(PfxFcmd$, CdlCmd$): ShellMax FcmdCrt(PfxFcmd, CdlCmd):      End Sub
Sub ShellFfnWin(FfnWin$):           ShellHid FmtQQ("cmd /c ""?""", FfnWin): End Sub
Sub ShellHid(StrCmd$):              Shell StrCmd, vbHide:                   End Sub
Sub ShellMax(StrCmd$):              Shell StrCmd, vbMaximizedFocus:         End Sub

Sub ChkBet(Fun$, V, VFm, VTo)
If VFm > V Then Thw Fun, "VFm > V", "V VFm VTo", V, VFm, VTo
If VTo < V Then Thw Fun, "VTo < V", "V VFm VTo", V, VFm, VTo
End Sub


Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function


Function Max(A, B, ParamArray Ap())
Dim O: O = IIf(A > B, A, B)
Dim J%: For J = 0 To UBound(Ap)
   If Ap(J) > O Then O = Ap(J)
Next
Max = O
End Function
Function Min(A, B, ParamArray Ap())
Dim O: O = IIf(A > B, B, A)
Dim J%: For J = 0 To UBound(Ap)
   If Ap(J) < O Then O = Ap(J)
Next
Min = O
End Function

Function CvVbt(A) As VbVarType: CvVbt = A: End Function

Function VbCprMth(C As eCas) As VbCompareMethod
Const CSub$ = CMod & "CprMth"
Select Case True
Case C = eCasSen: VbCprMth = vbBinaryCompare
Case C = eCasIgn: VbCprMth = vbTextCompare
Case Else: Thw CSub, "Invalid value of eCas", "eCas", C
End Select
End Function
Function CanCvLng(V) As Boolean
On Error GoTo X
Dim L&: L = CLng(V)
CanCvLng = True
X:
End Function

Sub SndKeys(A$)
DoEvents
SendKeys A, True
End Sub

Function NDig&(N)
Dim A$: A = N
NDig = Len(A)
End Function

Private Sub B_LngySum()
Dim S$: S = LinesFt(CPjf)
Dim Cnt&(): Cnt = AscCnty(S)
Debug.Assert LngySum(Cnt) = Len(S)
End Sub
Function LngySum@(A&())
Dim L: For Each L In Itr(A)
    LngySum = LngySum + L
Next
End Function

Function NBlk&(N&, SiBlk%)
NBlk = ((N - 1) \ SiBlk) + 1
End Function
Function CvDbl(S, Optional Fun$)
Const CSub$ = CMod & "CvDbl"
'Ret : a dbl of @S if can be converted, otherwise empty and debug.print S$
On Error GoTo X
CvDbl = CDbl(S)
Exit Function
X: If Fun <> "" Then Inf CSub, "str[" & S & "] cannot cv to dbl, emp is ret"
End Function

Function TmpnSno$(Optional Pfx$ = "N")
Static X&
TmpnSno = Tmpn(Pfx) & "_" & X
End Function
Function Tmpn$(Optional Pfx$ = "N"): Tmpn = Pfx & StrDte15(Now): End Function
