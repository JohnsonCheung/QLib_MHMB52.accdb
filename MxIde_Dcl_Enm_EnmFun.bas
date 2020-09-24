Attribute VB_Name = "MxIde_Dcl_Enm_EnmFun"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Enm_TEnm."
Type TEmbr: Mbn As String: Enmv As Long: Rmkl As String: End Type ' Deriving(Ctor Ay)
Type TEnm: IsPrv As Boolean: Enmn As String: Mbr() As TEmbr: IsGenEnmm4Fun As Boolean: IsGenEnmt As Boolean: End Type 'Deriving(Ctor Ay)
Private Sub B_EnmnyPC():      Brw EnmnyDcl(DclPC):        End Sub
Private Sub B_EnmblklsyDcl(): BrwLsy EnmblklsyDcl(DclPC): End Sub
Private Function B_EnmblkEnmnMC(): D EnmblkEnmnMC("AA1234"): End Function
Private Sub B_EnmblkyP(): BrwLyy EnmblkyP(CPj): End Sub
Function NEnmDcl%(Dcl$())
Dim L, O%
For Each L In Itr(Dcl)
   If IsLnEnm(L) Then O = O + 1
Next
NEnmDcl = O
End Function
Function NEnmM%(M As CodeModule): NEnmM = NEnmDcl(DclM(M)): End Function
Function TEnmyDcl(Dcl$(), Optional Mdn$) As TEnm()
Dim Blky(): Blky = EnmblkEnmny(Dcl)
Dim Blk: For Each Blk In Itr(Blky)
    PushTEnm TEnmyDcl, TEnmBlk(Blk, Mdn)
Next
End Function
Function MbnyTEnm(U As TEnm) As String()
Dim M() As TEmbr: M = U.Mbr
Dim J%: For J = 0 To UbTEmbr(M)
    PushI MbnyTEnm, M(J).Mbn
Next
End Function
Private Function EnmnyDcl(Dcl$()) As String()
Dim L: For Each L In Itr(Dcl)
    PushNB EnmnyDcl, EnmnLn(L)
Next
End Function
Function EnmnyM(M As CodeModule) As String():   EnmnyM = EnmnyDcl(DclM(M)): End Function
Function EnmnyMC() As String():                EnmnyMC = EnmnyDcl(DclMC):   End Function
Function EnmnyP(P As VBProject) As String():    EnmnyP = EnmnyDcl(DclP(P)): End Function
Function EnmnyPC() As String():                EnmnyPC = EnmnyP(CPj):       End Function
Function MenmnyPC() As String():              MenmnyPC = MenmnyP(CPj):      End Function
Function MenmnyP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MenmnyP, MenmnyCmp(C)
Next
End Function
Function MenmnyCmp(C As VBComponent) As String(): MenmnyCmp = AmAddPfx(EnmnyDcl(DclCmp(C)), C.Name & "."): End Function

Function EnmblkEnmn(Dcl$(), Enmn) As String(): EnmblkEnmn = AwBei(Dcl, X_Bei(Dcl, WEnmbixEnmn(Dcl, Enmn))): End Function
Function EnmblklsyDcl(Dcl$()) As String()
Dim B() As Bei: B = WEnmbeiyDcl(Dcl)
Dim J%: For J = 0 To UbBei(B)
    PushI EnmblklsyDcl, JnCrLf(AwBei(Dcl, B(J)))
Next
End Function
Function EnmblkEnmnM(M As CodeModule, EnmnLn) As String():  EnmblkEnmnM = EnmblkEnmn(DclM(M), EnmnLn):    End Function
Function EnmblkEnmnMC(EnmnLn) As String():                 EnmblkEnmnMC = EnmblkEnmnM(CMd, EnmnLn):       End Function
Function EnmblkEnmny(Dcl$()) As Variant():                  EnmblkEnmny = AyyBeiy(Dcl, WEnmbeiyDcl(Dcl)): End Function
Function EnmblkyM(M As CodeModule) As Variant():               EnmblkyM = EnmblkEnmny(DclM(M)):           End Function
Function EnmblkyMC() As Variant():                            EnmblkyMC = EnmblkyM(CMd):                  End Function
Function EnmblkyP(P As VBProject) As Variant()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy EnmblkyP, EnmblkyM(C.CodeModule)
Next
End Function
Private Function WEnmbeiyDcl(Dcl$()) As Bei()
Dim B&(): B = WBixy(Dcl)
Dim J%: For J = 0 To UB(B)
    PushBei WEnmbeiyDcl, X_Bei(Dcl, B(J))
Next
End Function
Private Function WEnmbixEnmn%(Dcl$(), EnmnLn_)
Dim O%, L: For Each L In Itr(Dcl)
    If EnmnLn(L) = EnmnLn_ Then
        WEnmbixEnmn = O
        Exit Function
    End If
    O = O + 1
Next
WEnmbixEnmn = -1
End Function
Private Function WBixy(Dcl$()) As Long()
Dim J&: For J = 0 To UB(Dcl)
    If IsLnEnm(Dcl(J)) Then PushI WBixy, J
Next
End Function
Private Function IsLnEnm(L) As Boolean:     IsLnEnm = HasPfx(RmvMdy(L), "Enum "):  End Function
Private Function X_Bei(Dcl$(), Bix) As Bei:   X_Bei = Bei(Bix, X_Eix(Dcl, Bix)):   End Function
Private Function X_Eix%(Dcl$(), Bix):         X_Eix = EixSrcItm(Dcl, Bix, "Enum"): End Function
Function TEmbrAdd(A As TEmbr, B As TEmbr) As TEmbr(): PushTEmbr TEmbrAdd, A: PushTEmbr TEmbrAdd, B: End Function
Sub PushTEmbry(O() As TEmbr, A() As TEmbr): Dim J&: For J = 0 To UbTEmbr(A): PushTEmbr O, A(J): Next: End Sub
Sub PushTEmbr(O() As TEmbr, M As TEmbr): Dim N&: N = SiTEmbr(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTEmbr&(A() As TEmbr): On Error Resume Next: SiTEmbr = UBound(A) + 1: End Function
Function UbTEmbr&(A() As TEmbr): UbTEmbr = SiTEmbr(A) - 1: End Function
Function TEmbr(Mbn, Enmv, Rmkl) As TEmbr
With TEmbr
    .Mbn = Mbn
    .Enmv = Enmv
    .Rmkl = Rmkl
End With
End Function

Sub PushTEnm(O() As TEnm, M As TEnm): Dim N&: N = SiTEnm(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTEnm&(A() As TEnm): On Error Resume Next: SiTEnm = UBound(A) + 1: End Function
Function UbTEnm&(A() As TEnm): UbTEnm = SiTEnm(A) - 1: End Function
Function TEnm(IsPrv, Enmn, Mbr() As TEmbr, IsGenEnmm4Fun, IsGenEnmt) As TEnm
With TEnm
    .IsPrv = IsPrv
    .Enmn = Enmn
    .Mbr = Mbr
    .IsGenEnmm4Fun = IsGenEnmm4Fun
    .IsGenEnmt = IsGenEnmt
End With
End Function
Function EnmvEnms&(Enmsy$(), Enms, Enmn$)
EnmvEnms = IxEle(Enmsy, Enmn): If EnmvEnms = -1 Then Inf CSub, "Enms not found in Enmsy", "Enmn Enms Enmsy", Enmn, Enms, Enmsy
End Function
