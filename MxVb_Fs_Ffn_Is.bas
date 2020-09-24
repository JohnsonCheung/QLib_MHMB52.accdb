Attribute VB_Name = "MxVb_Fs_Ffn_Is"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_Is."
Public Const FbExt$ = ".accdb"
Public Const FbExt1$ = ".mdb"
Public Const FbaExt$ = ".accdb"
Public Const FxaExt$ = ".xlam"
Function IsFxa(Ffn) As Boolean: IsFxa = LCase(Ext(Ffn)) = FxaExt: End Function
Function IsFba(Ffn) As Boolean: IsFba = LCase(Ext(Ffn)) = FbaExt: End Function
Function IsPjf(Ffn) As Boolean
Select Case True
Case IsFxa(Ffn), IsFba(Ffn): IsPjf = True
End Select
End Function
Function IsFb(Ffn) As Boolean
Select Case LCase(Ext(Ffn))
Case FbExt, FbExt1: IsFb = True
End Select
End Function
Function IsFx(Ffn) As Boolean
Select Case LCase(Ext(Ffn))
Case ".xls", ".xlsm", ".xlsx": IsFx = True
End Select
End Function
Function IsXlsx(Fn) As Boolean:   IsXlsx = HasSfx(Fn, ".xlsx"): End Function
Function IsXlsFn(Fn) As Boolean: IsXlsFn = HasSfx(Fn, ".xls"):  End Function
