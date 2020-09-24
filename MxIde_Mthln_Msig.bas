Attribute VB_Name = "MxIde_Mthln_Msig"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_MSig_."
Private Sub B_MsigLn() ' :L #AA# lskdf
GoSub T1
'GoSub T2
GoSub T3
'GoSub Z
Exit Sub
Dim Act As Msig, Ept As Msig, Mthln
Dim Mthl$, Ix&
T0:
    Mthln = "Function MsigLn(Mthln) As Msig() '  ksdljf    #AA-BB#"
    Dim A0 As TArg: A0 = TArg(eArgmRef, "Mthln", DftTVt, "")
    Dim A() As TArg: PushTArg A, A0
    Ept = Msig("", "Fun", "MsigLn", A, DftTVt, "ksdljf", "#AA-BB#")
    GoTo Tst
Z:
    BrwAy MsigsyMC
    Return
T1:
    Mthl = "Private Sub MsigsyM__Tst()"
    GoTo Tst
T2:
    Mthl = "Function MsigsyzM(M As CodeModule) As String()"
    GoTo Tst
T3:
    Mthl = "Sub PushBlk(O() As SBlk, M As SBlk)"
    GoTo Tst
Tst:
    Act = MsigLn(Mthln)
    If Not IsEqMsig(Act, Ept) Then Stop
    Return
End Sub
Private Sub MsigLn__Tst()
GoSub T1
Exit Sub
Dim Ln, Act As Msig
T1:
    Ln = "Friend Property Get MMym() As Ym"
    GoTo Tst
Tst:
    Act = MsigLn(Ln)
    Stop
    Return
End Sub
Function MsigLn(Ln) As Msig
Const CSub$ = CMod & "MsigLn"
Dim L$
With BrkVmk(Ln)
    L = .S1
    Dim Rmk$: Rmk = .S2
End With
Dim Mdy$: Mdy = ShfShtMdy(L):
Dim MTy$: MTy = ShfShtMthTy(L): If MTy = "" Then Thw CSub, "Given Ln is invalid: No mth ty", "Ln", Ln
Dim Nm$: Nm = ShfNm(L)
Dim Tyc$: Tyc = ShfTyc(L)
Dim Pm$: Pm = ShfBetBkt(L)
If IsShfAs(L) Then
    Dim Vsfx$: Vsfx = ShfVsfx(L)
    Dim T As TVt: T = TVtVsfx(Vsfx)
    Dim M$:  M = Memn(Rmk)
End If
MsigLn = Msig(Mdy, MTy, Nm, WTArgy(Pm), T, Rmk, M)
End Function

Function ShfVsfx$(OStrAftDimnOrArgn$)
Dim S$: S = OStrAftDimnOrArgn
Dim O$: O = ShfTyc(S)
If O <> "" Then
    ShfVsfx = O & IIf(IsShfBkt(O), "()", "")
    Exit Function
End If
Dim Bkt$:
    If IsShfBkt(O) Then
        Bkt = "()"
    End If
    Dim DNm$: DNm = ShfDotn(O):
    ShfVsfx = Bkt & " As " & DNm
    If DNm = "" Then Stop

    ShfVsfx = Bkt
End Function

Private Function WTArgy(Mthpm$) As TArg()
Dim A$(): A = SplitCmaSpc(Mthpm)
Dim Arg: For Each Arg In Itr(A)
    PushTArg WTArgy, TArgArg(Arg)
Next
End Function
Private Sub B_MsigsyM(): BrwAy MsigsyMC: End Sub
Private Sub B_MsigsyP(): BrwAy MsigsyPC: End Sub

Function MsigsLn$(Mthln):                        MsigsLn = Msigs(MsigLn(Mthln)):      End Function
Function MsigsyMC() As String():                MsigsyMC = MsigsyM(CMd):              End Function
Function MsigsyM(M As CodeModule) As String():   MsigsyM = MsigsySrc(SrcM(M)):        End Function
Function MsigsyPC() As String():                MsigsyPC = MsigsyP(CPj):              End Function
Function MsigsyP(P As VBProject) As String():    MsigsyP = MsigsyMthlny(MthlnyP(P)):  End Function
Function MsigsySrc(Src$()) As String():        MsigsySrc = MsigsyMthlny(Mthlny(Src)): End Function
Function MsigsyMthlny(Mthlny$()) As String()
Dim L, J&: For Each L In Itr(Mthlny)
    PushI MsigsyMthlny, MsigsLn(L)
    J = J + 1
Next
End Function
Function Msigs$(S As Msig)
Dim P1$:   P1 = WSS_P1(S)
Dim Pm$:   Pm = WSS_Pm(S.Arg)
        Msigs = TmlAp(P1, Pm, S.Memn, S.Rmk)
End Function
Private Function WSS_P1$(S As Msig): WSS_P1 = WMthMdyChr(S.ShtMdy) & S.Mthn & WSS_Vssfx(S.ShtMdy, S.ShtTy, S.Vt): End Function
Private Function WMthMdyChr$(ShtMdy$)
Const CSub$ = CMod & "WMthMdyChr"
Dim O$
Select Case True
Case ShtMdy = ""
Case ShtMdy = "Prv": O = "."
Case ShtMdy = "Frd": O = ":"
Case Else: Thw CSub, "ShtMdy error", "ShtMdy", ShtMdy
End Select
WMthMdyChr = O
End Function
Private Function WSS_Vssfx$(ShtMdy$, ShtTy$, RetVty As TVt) '#Vssfx:Var-sht-sfx#
Stop 'If ShtTy <> "Sub" Then WSS_Vssfx = StrPfxIfNB(Vsfx(RetVty), ":")
End Function
Private Function WSS_Pm$(A() As TArg)
Dim O$(), J%: For J = 0 To UbTArg(A)
    PushI O, ShtArgTArg(A(J))
Next
WSS_Pm = QuoTm(JnSpc(O))
End Function

Private Sub B_MMsigsyPC():                    BrwTmly MMsigsyPC: End Sub
Function MMsigsyPC() As String(): MMsigsyPC = MMsigsyP(CPj):     End Function
Function MMsigsyP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MMsigsyP, MMsigsyM(C.CodeModule)
Next
End Function
Function MMsigsyMC() As String():                MMsigsyMC = MMsigsyM(CMd):                       End Function
Function MMsigsyM(M As CodeModule) As String():   MMsigsyM = MMsigsySrc(SrcM(M), Mdn(M)):         End Function
Function MMsigsySrc(Src$(), Mdn$) As String():  MMsigsySrc = AmAddPfx(MsigsySrc(Src), Mdn & " "): End Function
