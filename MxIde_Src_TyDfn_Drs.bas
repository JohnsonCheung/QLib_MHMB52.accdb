Attribute VB_Name = "MxIde_Src_TyDfn_Drs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_TyDfn_Drs."
Public Const FfTTyDfn$ = "Mdn Nm Ty Mem Rmk"
Private Type X_TLn1: TyDfnn As String: Tyn As String: Memn As String: Rmk As String: End Type

Private Sub B_DrsTTyDfnPC():                              BrwDrs DrsTTyDfnPC:       End Sub
Function DrsTTyDfnPC() As Drs:              DrsTTyDfnPC = DrsTTyDfnP(CPj):          End Function ' Drs-Mdn-Nm-Ty-Mem-Rmk
Function DrsTTyDfnP(P As VBProject) As Drs:  DrsTTyDfnP = DrsFf(FfTTyDfn, W2Dy(P)): End Function
Private Function W2Dy(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    Dim Dy(): Dy = X_Dy(C.CodeModule)
    PushIAy W2Dy, Dy
Next
End Function

Function DrsTTyDfnMC() As Drs:               DrsTTyDfnMC = DrsTTyDfnM(CMd):          End Function
Function DrsTTyDfnM(M As CodeModule) As Drs:  DrsTTyDfnM = DrsFf(FfTTyDfn, X_Dy(M)): End Function

Private Function X_Dy(M As CodeModule) As Variant()
':X_Dy: :Dyo-Nm-Ty-Mem-LsyVmk ! Fst-Ln must be :nn: :dd #mm# !rr
'                                ! Rst-Ln is !rr
'                                ! must term: nn dd mm, all of them has no spc
'                                ! opt      : mm rr
'                                ! :xx:     : should uniq in pj
Dim Lsy$(): Lsy = LsyVmk(SrcM(M))
Dim N$: N = Mdn(M)
Dim R: For Each R In Itr(Lsy)
    PushI X_Dy, WDr(CvSy(R), N)
Next
End Function
Private Function WDr(LyTyDfn$(), Mdn$) As Variant()
Dim Dr(4): Stop 'Dr = WTLn(LyTyDfn(0))
Stop 'WDr(4) = AyAdd(Dr(4), LyRmk(CvSy(AeFst(LyTyDfn))))
WDr = Dr
End Function
Private Sub B_WTLn()
GoSub T1
Dim TLnTyDfn$, Act As X_TLn1, Ept As X_TLn1
Exit Sub
T1:
    TLnTyDfn = "':Cell: :SCell-or lsdf ldkjf "
    
Tst:
    Stop 'Act = WTLn(TLnTyDfn)
    Ass WIsEqTLn(Act, Ept)
    Return
    Return
End Sub
Private Function WTLn1(TLnTyDfn$) As X_TLn1
Const CSub$ = CMod & "WTLn1"
Dim L$: L = TLnTyDfn
Dim TyDfnn$, Tyn$, Memn$, Rmk$
TyDfnn = TyDfnnShf(L)
If TyDfnn = "" Then Thw CSub, "Given @TLnTyDfn does not have TyDfnn", "@TLnTyDfn", TLnTyDfn
Tyn = ColonTyShf(L)
Memn = MemnShf(L)
With WTLn1
    .Memn = Memn
    .Rmk = L
    .TyDfnn = TyDfnn
    .Tyn = Tyn
End With
End Function
Private Function WIsEqTLn(A As X_TLn1, B As X_TLn1) As Boolean
With A
    Select Case True
    Case .Memn <> B.Memn, .Rmk <> B.Rmk, .TyDfnn <> B.TyDfnn, .Tyn <> B.Tyn
    Case Else: WIsEqTLn = True
    End Select
End With
End Function
