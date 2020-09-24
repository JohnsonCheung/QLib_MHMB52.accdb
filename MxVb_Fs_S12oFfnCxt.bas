Attribute VB_Name = "MxVb_Fs_S12oFfnCxt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_S12."

Function S12yoFfnCxtzPth(Pth$, Optional Spec$ = "*.txt") As S12()
Dim P$: P = PthEnsSfx(Pth)
Dim F: For Each F In Itr(Fnay(Pth, Spec))
    Dim Ffn$: Ffn = P & F
    PushS12 S12yoFfnCxtzPth, S12oFfnCxt(Ffn)
Next
End Function

Function S12oFfnCxt(Ffn) As S12: S12oFfnCxt = S12(Ffn, LinesFt(Ffn)): End Function
