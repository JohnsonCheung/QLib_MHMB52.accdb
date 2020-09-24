Attribute VB_Name = "MxVb_Fs_OpInst"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_OpInst."

Function FfnInst$(Ffn):                 FfnInst = PthInst(Pth(Ffn)) & Fn(Ffn):          End Function
Function PthInstFdr$(Fdr):           PthInstFdr = PthAddFdrEns(PthTmpFdr(Fdr), StrNow): End Function
Function PthInstEns$(Pth):           PthInstEns = PthInst(Pth):                         End Function
Function IsFfnInst(Ffn) As Boolean:   IsFfnInst = IsFdrInst(FdrFfn(Ffn)):               End Function
Function IsFdrInst(Fdr$) As Boolean:  IsFdrInst = IsTimStr(Fdr):                        End Function
Function PthInst$(Pth):
Dim P$: P = PthEnsSfx(Pth)
PthInst = PthEns(P & FdrNxt(P) & "\")
End Function
