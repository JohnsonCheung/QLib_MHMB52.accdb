Attribute VB_Name = "MxIde_Mth_TMMth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_TMMth."
Type TIxMth: Ix As Long: Mthy() As String: Msig As Msig: End Type 'Deriving(Ctor Ay Opt)
Type TIxMthOpt: Som As Boolean: TIxMth As TIxMth: End Type
Type TMMth: Mdn As String: Mth() As TIxMth: End Type 'Deriving(Ctor Ay)
Type TPMth: Pjn As String: MMthAy() As TMMth: End Type 'Deriving(Ctor Ay)
Function TPMth(Pjn, MMthAy() As TMMth) As TPMth
With TPMth
    .Pjn = Pjn
    .MMthAy = MMthAy
End With
End Function
Function TPMthAdd(A As TPMth, B As TPMth) As TPMth(): PushTPMth TPMthAdd, A: PushTPMth TPMthAdd, B: End Function
Sub PushTPMthAy(O() As TPMth, A() As TPMth): Dim J&: For J = 0 To UbTPMth(A): PushTPMth O, A(J): Next: End Sub
Sub PushTPMth(O() As TPMth, M As TPMth): Dim N&: N = SiTPMth(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTPMth&(A() As TPMth): On Error Resume Next: SiTPMth = UBound(A) + 1: End Function
Function UbTPMth&(A() As TPMth): UbTPMth = SiTPMth(A) - 1: End Function
Function TMMth(Mdn, Mth() As TIxMth) As TMMth
With TMMth
    .Mdn = Mdn
    .Mth = Mth
End With
End Function
Function TMMthAdd(A As TMMth, B As TMMth) As TMMth(): PushTMMth TMMthAdd, A: PushTMMth TMMthAdd, B: End Function
Sub PushTMMthy(O() As TMMth, A() As TMMth): Dim J&: For J = 0 To UbTMMth(A): PushTMMth O, A(J): Next: End Sub
Sub PushTMMth(O() As TMMth, M As TMMth): Dim N&: N = SiTMMth(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTMMth&(A() As TMMth): On Error Resume Next: SiTMMth = UBound(A) + 1: End Function
Function UbTMMth&(A() As TMMth): UbTMMth = SiTMMth(A) - 1: End Function
Function TIxMth(Ix, Mthy$(), Msig As Msig) As TIxMth
With TIxMth
    .Ix = Ix
    .Mthy = Mthy
    .Msig = Msig
End With
End Function
Function TIxMthAdd(A As TIxMth, B As TIxMth) As TIxMth(): PushTIxMth TIxMthAdd, A: PushTIxMth TIxMthAdd, B: End Function
Sub PushTIxMthy(O() As TIxMth, A() As TIxMth): Dim J&: For J = 0 To UbTIxMth(A): PushTIxMth O, A(J): Next: End Sub
Sub PushTIxMth(O() As TIxMth, M As TIxMth): Dim N&: N = SiTIxMth(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTIxMth&(A() As TIxMth): On Error Resume Next: SiTIxMth = UBound(A) + 1: End Function
Function UbTIxMth&(A() As TIxMth): UbTIxMth = SiTIxMth(A) - 1: End Function
Function TIxMthOpt(Som, A As TIxMth) As TIxMthOpt: With TIxMthOpt: .Som = Som: .TIxMth = A: End With: End Function
Function SomTIxMth(A As TIxMth) As TIxMthOpt: SomTIxMth.Som = True: SomTIxMth.TIxMth = A: End Function

Function TMthySrc(Src$()) As TMth()
Dim L: For Each L In Itr(MthlnySrc(Src))
    PushTMth TMthySrc, TMthL(L)
Next
End Function

Private Sub B_TMthL()
GoSub T1
Exit Sub
Dim Ln$, Act As TMth, Ept As TMth
T1:
    Ln = "Function MthnnPubM$(M As CodeModule): MthnnPubM = JnSpc(AySrtQ(MthnPubyM(M, Patn))): End Function"
    GoTo Tst
Tst:
    Act = TMthL(Ln)
    Stop
    Return
End Sub
Function TMthL(Ln) As TMth:     TMthL = ShfTMth(CStr(Ln)):         End Function
Private Sub B_TMthyPC():                VcAy Mth3nyTMthy(TMthyPC): End Sub
Function TMthyPC() As TMth(): TMthyPC = TMthyP(CPj):               End Function
Function TMthyP(P As VBProject) As TMth()
Dim C As VBComponent: For Each C In P.VBComponents
    PushTMthy TMthyP, TMthyM(C.CodeModule)
Next
End Function
Function TMthyM(M As CodeModule) As TMth()
Dim L: For Each L In Itr(MthlnySrc(SrcM(M)))
    PushTMth TMthyM, TMthL(L)
Next
End Function

Function ShfTMth(OLn$) As TMth
Dim M$: M = ShfShtMdy(OLn)
Dim T$: T = ShfShtMthTy(OLn):: If T = "" Then Exit Function
ShfTMth = TMth(ShfNm(OLn), T, M)
End Function

Function RmvMth$(Ln)
Const CSub$ = CMod & "TMth_LnRmv"
Dim L$: L = Ln
RmvMdy L
If ShfMtht(L) = "" Then Exit Function
If ShfNm(L) = "" Then Thw CSub, "Not as SrcLin", "Ln", Ln
RmvMth = L
End Function

Sub DmpTMth(N As TMth):                         D Mi3NtfTMth(N):                                   End Sub
Function Mi3NtfTMth$(M As TMth):   Mi3NtfTMth = JnSpcAp(M.Mthn, M.ShtTy, StrDft(M.ShtMdy, "Pub")): End Function
Function Mit3NtfTMth$(M As TMth): Mit3NtfTMth = JnTabAp(M.Mthn, M.ShtTy, StrDft(M.ShtMdy, "Pub")): End Function
Function TMthMinus(A() As TMth, B() As TMth) As TMth()
If SiTMth(A) = 0 Then TMthMinus = B: Exit Function
If SiTMth(B) = 0 Then TMthMinus = A: Exit Function
Dim J%: For J = 0 To UbTMth(A)
    If Not HasTMth(B, A(J)) Then PushTMth TMthMinus, A(J)
Next
End Function
Function HasTMth(A() As TMth, M As TMth) As Boolean
Dim J%: For J = 0 To UbTMth(A)
    If IsEqTMth(A(J), M) Then HasTMth = True
Next
End Function
Private Sub B_Mi3NtfTMth(): MsgBox "[" & Mi3NtfTMth(TMthL("Function MthnTMth$(Mi3NtfTMth$)")) & "]": End Sub
Function IsEqTMth(A As TMth, B As TMth) As Boolean
With A
    Select Case True
    Case .Mthn <> B.Mthn, .ShtTy <> B.ShtTy, .ShtMdy <> B.ShtMdy
    Case Else: IsEqTMth = True
    End Select
End With
End Function
