Attribute VB_Name = "MxIde_Src_LisMth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_LisMth."


Sub LisPubFunRetAs(RetAsPatn$)
Dim RetSfx As Drs: Stop 'RetSfx = AddDcTMthRetTyn(DrsTMthCacFunPC)
Dim Patn As Drs: Patn = DwPatn(RetSfx, "RetSfx", RetAsPatn)
Dim T50 As Drs: T50 = DwTopN(Patn)
BrwDrs T50
End Sub
Sub VcMthWh(W As WhMth):                      WLisMth CPj, W, OupTy:=eOupVc, Top:=0: End Sub
Sub DmpMthWh(W As WhMth, Optional Top% = 50): WLisMth CPj, W, eOupDmp, Top:          End Sub
Private Sub WLisMth(P As VBProject, W As WhMth, OupTy As eOup, Top%)
Dim D As Drs
    Stop 'D = DrsTMthW(DrsTMthLisP(CPj), W)
    D = DwTopN(D, Top)
Dim O$()
    O = FmtDrs(D)
OupAy O, OupTy
End Sub
