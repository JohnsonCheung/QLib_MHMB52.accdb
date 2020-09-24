Attribute VB_Name = "MxIde_Mthln_Msig_Ud"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_MSig_Ud."
#If Doc Then
#End If
Enum eArgm: eArgmRef: eArgmVal: eArgmValOpt: eArgmRefOpt: eArgmAp: End Enum: Public Const EnmmArgm$ = "eArgm? Ref Val ValOpt RefOpt Ap"
Type TArg: Argm As eArgm: Argn As String: Vt As TVt: Dft As String: End Type 'Deriving(Ay Ctor) OptMbr(Dft)
Type Msig
    ShtMdy As String
    ShtTy As String  ' ShtMthTy
    Mthn As String
    Arg() As TArg
    Vt As TVt
    Rmk As String ' On the same lines
    Memn As String
End Type ' Deriving(Ay Ctor)
Function TArg(Argm As eArgm, Argn, Vt As TVt, Dft) As TArg
With TArg
    .Argm = Argm
    .Argn = Argn
    .Vt = Vt
    .Dft = Dft
End With
End Function
Function eArgm$(A As Long)
Stop
End Function
Function IsEqMsig(A As Msig, B As Msig) As Boolean
With A
    Select Case True
    Case .Mthn <> B.Mthn
    Case .ShtMdy <> B.ShtMdy
    Case .Rmk <> B.Rmk
    Case .ShtTy <> B.ShtTy
    Case IsEqTArgy(.Arg, B.Arg)
    Case Else: IsEqMsig = True
    End Select
End With
End Function
Function IsEqTArgy(A() As TArg, B() As TArg) As Boolean
Dim U&: U = UbTArg(A): If U <> UbTArg(B) Then Exit Function
Dim J%: For J = 0 To U
    If IsEqTArg(A(J), B(J)) Then Exit Function
Next
IsEqTArgy = True
End Function
Function IsEqTArg(A As TArg, B As TArg) As Boolean
With A
    Select Case True
    Case .Argm <> B.Argm
    Case .Argn <> B.Argn
    Case .Dft <> B.Dft
    Case Not IsEqVt(.Vt, B.Vt)
    Case Else: IsEqTArg = True
    End Select
End With
End Function
Function IsEqVt(A As TVt, B As TVt) As Boolean
With A
    Select Case True
    Case .IsAy <> B.IsAy
    Case .Tyc <> B.Tyc
    Case .Tyn <> B.Tyn
    Case Else: IsEqVt = True
    End Select
End With
End Function

Function VsfxTVt$(T As TVt) ' #Variable-Sfx# short variable suffix directly added to varn which can be a dimn argn udtn
Dim B$: B = BktIf(T.IsAy)
Select Case True
Case T.Tyn = "" And T.Tyc = "": VsfxTVt = B
Case T.Tyn = "":                VsfxTVt = T.Tyc & B
Case Else
    Dim C$: C = TycTycn(T.Tyn)
    If C = "" Then
        VsfxTVt = ":" & Shttyn(T.Tyn) & B
    Else
        VsfxTVt = C & B
    End If
End Select
End Function
Function Argmy() As String()
Const T0 = "ByRef"
Const T1 = "ByVal"
Const T2 = "Optional ByVal"
Const T3 = "Optional ByRef"
Const T4 = "ParamArray"
Static X$(): If Si(X) = 0 Then X = Sy(T0, T1, T2, T3, T4)
Argmy = X
End Function
Function DftTVt() As TVt: DftTVt = TVt("", False, ""): End Function
Sub PushTVt(O() As TVt, A() As TVt): Dim J&: For J = 0 To UbTVt(A): PushTVty O, A(J): Next: End Sub
Sub PushTVty(O() As TVt, M As TVt): Dim N&: N = SiTVt(O): ReDim Preserve O(N): O(N) = M: End Sub
Function TArgAdd(A As TArg, B As TArg) As TArg(): PushTArg TArgAdd, A: PushTArg TArgAdd, B: End Function
Sub PushTArgy(O() As TArg, A() As TArg): Dim J&: For J = 0 To UbTArg(A): PushTArg O, A(J): Next: End Sub
Sub PushTArg(O() As TArg, M As TArg): Dim N&: N = SiTArg(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTArg&(A() As TArg): On Error Resume Next: SiTArg = UBound(A) + 1: End Function
Function UbTArg&(A() As TArg): UbTArg = SiTArg(A) - 1: End Function
Function Msig(ShtMdy, ShtTy, Mthn, Arg() As TArg, Vt As TVt, Rmk, Memn) As Msig
With Msig
    .ShtMdy = ShtMdy
    .ShtTy = ShtTy
    .Mthn = Mthn
    .Arg = Arg
    .Vt = Vt
    .Rmk = Rmk
    .Memn = Memn
End With
End Function
Function AddMsig(A As Msig, B As Msig) As Msig(): PushMsig AddMsig, A: PushMsig AddMsig, B: End Function
Sub PushMsigAy(O() As Msig, A() As Msig): Dim J&: For J = 0 To MsigUB(A): PushMsig O, A(J): Next: End Sub
Sub PushMsig(O() As Msig, M As Msig): Dim N&: N = MsigSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function MsigSI&(A() As Msig): On Error Resume Next: MsigSI = UBound(A) + 1: End Function
Function MsigUB&(A() As Msig): MsigUB = MsigSI(A) - 1: End Function
