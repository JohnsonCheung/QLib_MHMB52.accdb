VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcLas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_Fc."
Private Fun As New MxMHMB52_MHFun
Function StmYmUD() As StmYm:         StmYmUD = StmYm("U"):                  End Function
Function StmYmMH() As StmYm:         StmYmMH = StmYm("M"):                  End Function
Function StmYmUDPthi() As StmYm: StmYmUDPthi = StmYmPthi("U"):              End Function
Function StmYmMHPthi() As StmYm: StmYmMHPthi = StmYmPthi("M"):              End Function
Function FxiMH$():                     FxiMH = MH.FcIO.Fxi(StmYmPthi("M")): End Function
Function FxiUD$():                     FxiUD = MH.FcIO.Fxi(StmYmPthi("U")): End Function

Function StmYm(Stm$) As StmYm
Dim Ym$: Ym = ValQC("Select Max(Format(VerYY,'00') & Format(VerMM,'00')) As YM from Fc where Stm='" & Stm & "'")
With StmYm
    .Stm = Stm
    .Y = Left(Ym, 2)
    .M = Right(Ym, 2)
End With
End Function
Function Ym() As Ym
Dim A$: A = ValQC("Select Max(Format(VerYY,'00') & Format(VerMM,'00')) As YM from Fc")
With Ym
    .Y = Left(A, 2)
    .M = Right(A, 2)
End With
End Function

Function FxiUDPthi$():             FxiUDPthi = EleMax(WFfny("U")):       End Function
Function FxiMHPthi$():             FxiMHPthi = EleMax(WFfny("M")):       End Function
Function StmYmPthi(Stm$) As StmYm: StmYmPthi = WStmYmzFn(Fn(FxiUDPthi)): End Function
Private Function WStmYmzFn(FnFc) As StmYm
With WStmYmzFn
    .Stm = Fun.StmStm2(WStm2(FnFc))
    .Y = Wyy(FnFc)
    .M = Mid(FnFc, 9, 2)
End With
End Function
Private Function WFfny(Stm$) As String()
Dim P$: P = MH.FcIO.Pthi
Dim F$(): F = Fnay(P, "Fc 20??-?? ??.xlsx")
Dim S2$: S2 = Fun.Stm2Stm(Stm)
Dim Fn: For Each Fn In Itr(F)
    If WIsFnFc(Fn, S2) Then PushI WFfny, P & Fn
Next
End Function
Private Function WIsFnFc(Fn, Stm2$) As Boolean
'123456789012345678
'Fc 20??-?? ??.xlsx
Select Case True
Case Len(Fn) <> 18
Case Left(Fn, 5) <> "Fc 20"
Case Not IsNumeric(Wyy(Fn))
Case Ext(Fn) <> ".xlsx"
Case Mid(Fn, 8, 1) <> "-"
Case Mid(Fn, 11, 1) <> " "
Case Not IsMM(Mid(Fn, 9, 2))
Case WStm2(Fn) <> Stm2
Case Else
    WIsFnFc = True
End Select
End Function
Private Function WStm2$(FnFc): WStm2 = Mid(FnFc, 12, 2): End Function
Private Function Wyy$(FnFc):     Wyy = Mid(FnFc, 6, 2):  End Function
