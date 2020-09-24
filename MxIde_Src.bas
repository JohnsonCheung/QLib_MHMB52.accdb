Attribute VB_Name = "MxIde_Src"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src."
Function SrcVC() As String(): SrcVC = SrcV(CVbe): End Function
Function SrcPC() As String(): SrcPC = SrcP(CPj):  End Function
Function SrcMC() As String(): SrcMC = SrcM(CMd):  End Function
Function SrcV(V As VBE) As String()
Dim P As VBProject: For Each P In V.VBProjects
    PushIAy SrcV, SrcP(P)
Next
End Function
Function SrcP(P As VBProject) As String()
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy SrcP, SrcM(C.CodeModule)
Next
End Function
Function SrcyP(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI SrcyP, SrcM(C.CodeModule)
Next
End Function
Function SrcM(M As CodeModule) As String()
Dim A$: A = SrclM(M): If A = "" Then Exit Function
SrcM = SplitCrLf(A)
End Function

Function SrclM$(M As CodeModule)
If M.CountOfLines = 0 Then Exit Function
SrclM = M.Lines(1, M.CountOfLines)
End Function

Function SrclsyP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI SrclsyP, SrclM(C.CodeModule)
Next
End Function
Function SrclsyPC() As String(): SrclsyPC = SrclsyP(CPj): End Function

Function SrclMdnPC$(Mdn):                SrclMdnPC = SrclMdnP(CPj, Mdn):    End Function
Function SrclMdnP$(P As VBProject, Mdn):  SrclMdnP = SrclCmp(CmpP(P, Mdn)): End Function

Function SrclMC$():                   SrclMC = SrclM(CMd):        End Function
Function SrclPC$():                   SrclPC = SrclP(CPj):        End Function
Function SrclP$(P As VBProject):       SrclP = JnCrLf(SrcP(P)):   End Function
Function SrclCmp$(C As VBComponent): SrclCmp = JnCrLf(SrcCmp(C)): End Function

Private Sub B_SrcRmvIfFalse(): Brw SrcRmvIfFalse(SrcMC): End Sub
Private Function SrcRmvIfFalse(Src$()) As String()
Const CSub$ = CMod & "SrcRmvIfFalse"
Dim InFalse As Boolean, IsFalseLn As Boolean, IsEndIfLn As Boolean
Dim L: For Each L In Itr(Src)
    IsFalseLn = L = "#If False Then"
    IsEndIfLn = L = "#End If"
    Select Case True
    Case IsFalseLn And InFalse:   Thw CSub, "Impossible to have InFalse=True and IsFalseLn=true"
    Case IsFalseLn:               InFalse = True
    Case IsEndIfLn And InFalse:   InFalse = False
    Case InFalse:
    Case Else:                    PushI SrcRmvIfFalse, L
    End Select
Next
End Function
