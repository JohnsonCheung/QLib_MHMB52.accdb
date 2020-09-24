Attribute VB_Name = "MxIde_Dcl"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl."
Public Const DclItmss$ = "Option Const Type Enum Dim "

Private Sub B_DclSrc()
GoSub ZZ
Exit Sub
Dim Src$()
ZZ:
    Src = SrcMdn("MxDao_Db_Schm_Ud")
    Act = DclSrc(Src)
    BrwAy Act
    Return
End Sub

Function HasSrclDcl(M As CodeModule, Srcl$) As Boolean
HasSrclDcl = HasSsub(DcllM(M), Srcl)
End Function

Function VarnyDimy(StmtyDim$()) As String()
Dim Stmt: For Each Stmt In Itr(StmtyDim)
    PushIAy VarnyDimy, VarnyDim(Stmt)
Next
End Function
Function VarnyDim(StmtDim) As String()
Dim L$: L = StmtDim
If Not IsShfPfx(L, "Dim ") Then Exit Function
VarnyDim = SplitCmaSpc(L)
End Function

Private Sub B_SrcRmvVmk(): Brw SrcRmvVmk(SrcP(CPj)): End Sub
Function IxCdlnPrv&(Src$(), Fm)
Dim O&: For O = Fm - 1 To 0 Step -1
    If IsLnCd(Src(O)) Then IxCdlnPrv = O: Exit Function
Next
IxCdlnPrv = -1
End Function
Function DiDclPC() As Dictionary
Set DiDclPC = DiDclP(CPj)
End Function

Function DiDclP(P As VBProject) As Dictionary
If P.Protection = vbext_pp_locked Then Set DiDclP = New Dictionary: Exit Function
Dim C As VBComponent, M As CodeModule
Set DiDclP = New Dictionary
For Each C In P.VBComponents
    Set M = C.CodeModule
    Dim D$(): D = DclM(M)
    If Si(D) = 0 Then
        DiDclP.Add MdnDotM(M), D
    End If
Next
End Function

Function DclPC() As String(): DclPC = DclP(CPj): End Function
Function DclP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy DclP, DclM(C.CodeModule)
Next
End Function

Function DcllSrc$(Src$()): DcllSrc = JnCrLf(DclSrc(Src)): End Function

Function HasDclln(M As CodeModule, Dclln$) As Boolean
Dim J&: For J = 1 To M.CountOfDeclarationLines
    If M.Lines(J, 1) = Dclln Then HasDclln = True: Exit Function
Next
End Function

Function DclSrc(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim N&: N = NDclln(Src)
If N <= 0 Then Exit Function
DclSrc = AwFstN(Src, N)
End Function
Private Sub B_Mdl()
Dim LinesBef$
Dim LinesAft$
Dim C As VBComponent: For Each C In CPj.VBComponents
    LinesBef = SrclCmp(C)
    LinesAft = LinesApLnNB(DcllCmp(C), BdylCmp(C))
    If LinesBef <> LinesAft Then Stop
Next
End Sub
Function BdylM$(M As CodeModule)
With M
Dim Lno&: Lno = .CountOfDeclarationLines
Dim Cnt&: Cnt = .CountOfLines - .CountOfDeclarationLines
End With
BdylM = M.Lines(Lno, Cnt)
End Function
Function BdylCmp$(C As VBComponent):           BdylCmp = BdylM(C.CodeModule):   End Function
Function BdyCmp(C As VBComponent) As String():  BdyCmp = SplitCrLf(BdylCmp(C)): End Function
Function Mdly(M As CodeModule) As String():       Mdly = SplitCrLf(BdylM(M)):   End Function
Function Mdlsy(P As VBProject, Mdny$()) As String()
Dim N: For Each N In Itr(Mdny)
    PushI Mdlsy, SrclMdnP(P, N)
Next
End Function

Function SrclLcntM$(M As CodeModule, Lc As Lcnt)
With Lc
If Lc.Lno <= 0 Then Exit Function
If Lc.Cnt <= 0 Then Exit Function
If Lc.Lno > M.CountOfLines Then Exit Function
SrclLcntM = M.Lines(Lc.Lno, Lc.Cnt)
End With
End Function

Private Sub B_Dcl()
Dim M As CodeModule
'GoSub ZZ
GoSub T1
Exit Sub
T1:
    Set M = Md("MxVbRunThw")
    Ept = ""
    GoTo Tst
Tst:
    Act = DclM(M)
    C
    Return
ZZ:
    Dim O$(), Cmp As VBComponent
    For Each Cmp In CPj.VBComponents
        PushNB O, DclM(Cmp.CodeModule)
    Next
    VcLsy O
End Sub
Function DclM(M As CodeModule) As String(): DclM = SplitCrLf(DcllM(M)): End Function
Function DclCmp(C As VBComponent) As String()
Dim M As CodeModule: Set M = MdCmp(C): If IsNothing(M) Then Exit Function
DclCmp = DclM(M)
End Function
Function DcllCmp$(C As VBComponent)
Dim M As CodeModule: Set M = MdCmp(C): If IsNothing(M) Then Exit Function
DcllCmp = DcllM(M)
End Function

Function DcllP()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushI O, DcllCmp(C)
Next
DcllP = JnCrLf(O)
End Function
Function DcllMC$(): DcllMC = DcllM(CMd): End Function
Function DcllM$(M As CodeModule)
Dim N&: N = M.CountOfDeclarationLines
If N > 0 Then DcllM = M.Lines(1, N)
End Function

Private Sub B_NDclln()
GoSub T1
Exit Sub
Dim Mdn$
T1:
    Mdn = "MxDaoDbLnk"
    Ept = 1
    GoTo Tst
Tst:
    Act = NDclln(SrcMdn(Mdn))
    Stop
    C
    Return
ZZ:
    Dim O$()
    Dim Cmp As VBComponent: For Each Cmp In CPj.VBComponents
        PushI O, NDclln(SrcCmp(Cmp)) & " " & Cmp.Name
    Next
    BrwAy O
    Return
End Sub
Function DclMC() As String(): DclMC = DclM(CMd): End Function
Function DclyP(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI DclyP, DclM(C.CodeModule)
Next
End Function
