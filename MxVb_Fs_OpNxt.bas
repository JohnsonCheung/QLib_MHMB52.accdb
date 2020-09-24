Attribute VB_Name = "MxVb_Fs_OpNxt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_OpNxt."

Private Sub B_PthNxtEns()
Dim P$: P = PthTmp
Dim J%: For J = 1 To 10
    PthNxtEns P
Next
D W1_FdryNxt(P)
End Sub
Function PthNxtEns$(Pth$): PthNxtEns = PthEns(PthNxt(Pth)): End Function
Function PthNxt$(Pth) ' It is a child pth of @Pth with Fdr being :FdrNxt
PthNxt = PthAddFdrEns(Pth, FdrNxt(Pth))
End Function
Function FdrNxt$(Pth):               FdrNxt = Pad0(W1_MaxFdr(Pth) + 1, 4): End Function
Private Function W1_MaxFdr%(Pth): W1_MaxFdr = EleMax(W1_FdryNxt(Pth)):     End Function

Private Function W1_FdryNxt(Pth) As String()
Dim A$(): A = Fdry(Pth, "????")
Dim I: For Each I In Itr(A)
    If W1_IsFdrNxt(I) Then PushI W1_FdryNxt, I
Next
End Function
Private Function W1_IsFdrNxt(S) As Boolean
Select Case True
Case Not IsStr(S), Len(S) <> 4, Not IsNumeric(S)
Case Else: W1_IsFdrNxt = True
End Select
End Function
