Attribute VB_Name = "MxVb_Fs_Pth_PthAdd"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Pth_OpVal."

Function PthAddFdr$(Pth, Fdr):       PthAddFdr = PthAddSeg(Pth, Fdr):                End Function
Function PthAddSeg$(Pth, SegPth):    PthAddSeg = PthEnsSfx(Pth) & PthEnsSfx(SegPth): End Function
Function PthAddFdrEns$(Pth, Fdr): PthAddFdrEns = PthEns(PthAddFdr(Pth, Fdr)):        End Function
Function PthAddFdrApEns$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
Dim O$: O = PthAddFdrAv(Pth, Av)
PthEnsAll O
PthAddFdrApEns = O
End Function

Function PthAddFdrAv$(Pth, FdrAv())
Dim O$: O = Pth
Dim I, Fdr$
For Each I In FdrAv
    Fdr = I
    O = PthAddFdr(O, Fdr)
Next
PthAddFdrAv = O
End Function

Function PthAddFdrAp$(Pth, ParamArray FdrAp())
Dim Av(): Av = FdrAp
PthAddFdrAp = PthAddFdrAv(Pth, Av)
End Function
