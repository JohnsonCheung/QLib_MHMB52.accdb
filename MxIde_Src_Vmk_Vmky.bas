Attribute VB_Name = "MxIde_Src_Vmk_Vmky"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Vmk_Vmky."
Private Sub B_VmkyFst(): BrwAy VmkyFst(SrcPC): End Sub

Function VmkyFm(Src$(), Fmix) As String()
Dim B&: B = BxVrmk(Src, Fmix): If B < 0 Then Exit Function
Dim E&: E = ExVmk(Src, B)
Dim O$(): O = AwBE(Src, B, E)
O(0) = Vmk(O(0))
VmkyFm = O
End Function

Private Function BxVrmk&(Src$(), Optional Fmix = 0)
Dim J&: For J = Fmix To UB(Src)
    Dim L$: L = Src(J)
    If HasQuoSng(L) Then
        If IsLnVmk(L) Then BxVrmk = J: Exit Function
        If Vmk(L) <> "" Then BxVrmk = J: Exit Function
    End If
Next
BxVrmk = -1
End Function
Private Function ExVmk&(Src$(), BxVmk&)
Dim IsContPrv As Boolean
    IsContPrv = ChrLas(Src(BxVmk)) = "_"
Dim J&: For J = BxVmk + 1 To UB(Src)
    Dim L$: L = Src(J)
    If Not IsLnVmkCont(L, IsContPrv) Then
        ExVmk = J - 1
        Exit Function
    End If
    IsContPrv = ChrLas(L) = "_"
Next
ExVmk = UB(Src)
End Function
Private Function IsLnVmkCont(L$, IsContPrv As Boolean) As Boolean
Select Case True
Case IsContPrv: IsLnVmkCont = True
Case Not HasQuoSng(L)
Case IsLnVmk(L), Vmk(L) <> "": IsLnVmkCont = True
End Select
End Function

Function VmklFm$(Src$(), Optional Fmix%): VmklFm = JnCrLf(VmkyFm(Src, Fmix)): End Function

Function VmkyFst(Src$()) As String(): VmkyFst = VmkyFm(Src, 0):       End Function
Function VmklFst$(Src$()):            VmklFst = JnCrLf(VmkyFst(Src)): End Function
