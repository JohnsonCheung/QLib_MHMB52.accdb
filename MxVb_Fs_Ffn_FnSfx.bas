Attribute VB_Name = "MxVb_Fs_Ffn_FnSfx"
Option Compare Text
Const CMod$ = "MxVb_Fs_Ffn_FnPfxSfx."
Option Explicit

Function FfnAddFnPfx$(Ffn, PfxFn$): FfnAddFnPfx = Pth(Ffn) & PfxFn & Fn(Ffn): End Function

Function FfnAddTimSfx$(Ffn):      FfnAddTimSfx = FfnAddFnsfx(Ffn, Format(Now, "(HHMMSS)")): End Function
Function FfnAddFnsfx$(Ffn, Sfx$):  FfnAddFnsfx = Ffnn(Ffn) & Sfx & Ext(Ffn):                End Function
Function FfnRmvFnsfx$(Ffn, Sfx$):  FfnRmvFnsfx = RmvSfx(Ffnn(Ffn), Sfx) & Ext(Ffn):         End Function
Function FfnyRmvFnsfx(Ffny$(), Sfx$) As String()
Dim Ffn: For Each Ffn In Itr(Ffny)
    PushI FfnyRmvFnsfx, FfnRmvFnsfx(Ffn, Sfx)
Next
End Function
Function FfnCutExtRmvFnsfx$(Ffn, Sfx$): FfnCutExtRmvFnsfx = RmvSfx(CutExt(Ffn), Sfx):               End Function
Function FfnRplFnsfx$(Ffn, Sfx$, By$):        FfnRplFnsfx = RmvSfx(Ffnn(Ffn), Sfx) & By & Ext(Ffn): End Function
Function FfnyCutExtRmvFnsfx(Ffny$(), Sfx$) As String()
Dim Ffn: For Each Ffn In Itr(Ffny)
    PushI FfnyCutExtRmvFnsfx, FfnCutExtRmvFnsfx(Ffn, Sfx)
Next
End Function
