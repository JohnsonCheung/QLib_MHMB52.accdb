Attribute VB_Name = "MxIde_Src_Cac_Ft_zIntl_ClrFtcacExcess"
Option Compare Text
Option Explicit
Public Const FnsfxFtcac$ = "(Ftcac)"
Public Const FnsfxFtcacMit8$ = "(Ftcac.Mit8Cmfntbel)"

Sub ClrFtcacP(P As VBProject):       ClrPth PthFtcacP(P):           End Sub
Sub ClrFtcacPC():                    ClrFtcacP CPj:                 End Sub
Sub ClrFtcacExcessP(P As VBProject): DltFfnyIf FfnyFtcacExcessP(P): End Sub
Sub ClrFtcacExcessPC():              ClrFtcacExcessP CPj:           End Sub
Private Function FfnyFtcacExcessP(P As VBProject) As String()
Dim Pth$: Pth = PthFtcacP(P)
Dim NyCur$(): NyCur = MdnySrcP(P)
Dim NyFtcac$(): NyFtcac = MdnyFtcac(Pth)
Dim NyExcess$(): NyExcess = SyMinus(NyFtcac, NyCur)
Dim N: For Each N In Itr(NyExcess)
    DltFfnIf Pth & N & FnsfxFtcac & ".txt"
    DltFfnIf Pth & N & FnsfxFtcacMit8 & ".txt"
Next
End Function
Private Function MdnyFtcac(Pth) As String()
Dim Fn: For Each Fn In Itr(Fnay(Pth, "*.txt"))
    PushNoDup MdnyFtcac, MdnFnFtcac(Fn)
Next
End Function
Private Function MdnFnFtcac$(Fn): MdnFnFtcac = RmvSfx(RmvSfx(FnnFfn(Fn), FnsfxFtcac), FnsfxFtcacMit8): End Function
