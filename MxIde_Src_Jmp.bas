Attribute VB_Name = "MxIde_Src_Jmp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Jmp."
Private Type MdnLcc
    Mdn As String
    Lno As Long
    C1 As Integer
    C2 As Integer
End Type

Private Sub B_Jmp():         Jmp "MxGit:1":                           End Sub
Sub JmpMdfssubSen(Mdfssub$): JmpMdny MdnyMdfssub(Mdfssub, eCasSen):   End Sub
Sub JmpMdssubSen(Mdssub$):   JmpMdny AwSsub(MdnyPC, Mdssub, eCasSen): End Sub
Sub JmpMdfssub(Mdfssub$):    JmpMdny MdnyMdfssub(Mdfssub):            End Sub
Sub JmpMdssub(Mdssub$):      JmpMdny AwSsub(MdnyPC, Mdssub):          End Sub
Sub JmpMdny(Mdny$())
Select Case Si(Mdny)
Case 1: JmpMdn Mdny(0)
Case Is > 1
    Dim J%: For J = 0 To UB(Mdny)
        Debug.Print Mdny(J)
    Next
End Select
End Sub
Sub JmpMxMHMB52():      JmpMdnPfx "MxMHMB52":  End Sub
Sub JmpMdnPfx(PfxMdn$): JmpMdn MdnPfx(PfxMdn): End Sub
Sub JmpMdn(Mdn):        JmpMd Md(Mdn):         End Sub
Sub JmpMdC():           JmpMd CMd:             End Sub
Sub JmpMdf(Mdf$):       JmpMd MdMdf(Mdf):      End Sub
Sub JmpMd(M As CodeModule)
If IsNothing(M) Then ClsWinAll: Exit Sub
If Not M.CodePane.Window.Visible Then M.CodePane.Show
M.CodePane.Show
VisIWin IWinPj
Dim W As VbIde.Window: For Each W In CVbe.Windows
    Select Case True
    Case _
        W.Type = vbext_wt_Immediate, _
        IsIWinEqMd(W, M)
'    Case _
        W.Type = vbext_wt_ProjectWindow, _
        W.Type = vbext_wt_Immediate, _
        IsIWinEqMd(W, M)
    Case Else
        ClsWin W
    End Select
Next
With M.CodePane.Window
    .SetFocus
    .WindowState = vbext_ws_Maximize
End With
DoEvents
ClsIWinImm
End Sub
Function IsIWinEqMd(W As VbIde.Window, M As CodeModule) As Boolean
If IsNothing(M) Then Exit Function
IsIWinEqMd = IsEqObj(W, M.CodePane.Window)
End Function
Sub JmpMdLno(M As CodeModule, Lno&)
JmpMd M
JmpLno Lno
End Sub

Function MdnLcczRep(RepMdnLcc$) As MdnLcc
Const CSub$ = CMod & "MdnLccRep"
Dim A$(): A = SplitColon(RepMdnLcc)
Dim N%: N = Si(A): If Not IsBet(N, 2, 4) Then Thw CSub, "N-term-of-MdnLccRep should between 2 to 4 Itm sep by [:]", "N-Term-of-RepMdnLcc", RepMdnLcc
With MdnLcczRep
    .Mdn = A(0)
    .Lno = A(1)
    If N >= 2 Then .C1 = A(2)
    If N >= 3 Then .C2 = A(3)
End With
End Function
Sub Jmp(RepMdnLcc$) 'RepMdnLcc: #Mdn-Lno-C1-C2-Str# Rep/Mdn:Lno:C1:C2/ where :C1:C2 is optional
With MdnLcczRep(RepMdnLcc)
If .Mdn <> "" Then JmpMdn .Mdn
If .Lno > 0 Then
    JmpLno .Lno
    If .C1 > 0 Then CPne.SetSelection .Lno, .C1, .Lno, .C2 + 1
End If
End With
End Sub

Sub JmpRCC(R&, C1%, C2%)
CPne.SetSelection R, C1, R, C2
End Sub

Sub JmpLno(Lno&)
Dim C2%: C2 = Len(CMd.Lines(Lno, 1)) + 1
JmpLcc Lno, 1, C2
End Sub

Sub JmpLcc(Lno&, C1%, C2%)
Dim L1&: L1 = Lno - 6: If L1 <= 0 Then L1 = 1
With CPne
    .TopLine = L1
    .SetSelection Lno, C1, Lno, C2
End With
End Sub

Sub JmpRrcc(A As Rrcc)
Dim L&, C1%, C2%
With CPne
    If C1 = 0 Or C2 = 0 Then
        C1 = 1
        C2 = Len(.CodeModule.Lines(L, 1)) + 1
    End If
    .TopLine = L
    .SetSelection L, C1, L, C2
End With
'SendKeys "^{F4}"
End Sub
Function MdnyMthn(Mthn) As String(): MdnyMthn = MdnyMthnP(CPj, Mthn): End Function
Function MdnyMthnP(P As VBProject, Mthn) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If HasMthnM(C.CodeModule, Mthn) Then
        PushI MdnyMthnP, C.Name
    End If
Next
End Function
Sub JmpMth(Mthn)
Dim N$(): N = MdnyMthnP(CPj, Mthn)
Select Case Si(N)
Case 0: Debug.Print "No mth [" & Mthn & "]"
Case 1: JmpMdnMth N(0), Mthn
Case Else
    Dim J%: For J = 0 To UB(N)
        Debug.Print FmtQQ("JmpMdnMth ""?"", ""?""", N(J), Mthn)
    Next
End Select
Dim M As CodeModule: Set M = Md(N(0))
JmpMd M
JmpLno Mthlno(M, Mthn)
End Sub

Sub JmpMdnMth(Mdn$, Mthn)
JmpMd Md(Mdn)
JmpMth Mthn
End Sub

Sub JmpPj(P As VBProject)
ClsWinAll
Dim M As CodeModule
Set M = MdFst(P)
If IsNothing(M) Then Exit Sub
JmpMd M
TileV
DoEvents
End Sub

Sub JmpMdRrcc(M As CodeModule, R As Rrcc)
JmpMd M
JmpRrcc R
End Sub

Sub JmpMdnn(Mdnn$)
ClsWinAll
Dim M: For Each M In Itr(SySs(Mdnn))
    VisMd Md(M)
Next
TileV
End Sub

Sub VisMd(M As CodeModule)
M.CodePane.Window.Visible = True
End Sub
