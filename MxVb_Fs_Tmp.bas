Attribute VB_Name = "MxVb_Fs_Tmp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Tmp."
Function FfnTmpCxt$(Ext, Cxt$, Optional Fdr$, Optional PfxFn$)
Dim Ffn$: Ffn = FfnTmp(Ext, Fdr, PfxFn)
WrtStr Cxt, Ffn
FfnTmpCxt = Ffn
End Function

Function FfnTmp$(Ext, Optional Fdr$, Optional PfxFn$):    FfnTmp = PthTmpFdr(Fdr) & Tmpn(StrDft(PfxFn, "N")) & Ext: End Function
Function FfnTmpCsv$(Optional Fdr$, Optional PfxFn$):   FfnTmpCsv = FfnTmp(".csv", Fdr, PfxFn):                      End Function
Function FtTmp$(Optional Fdr$, Optional Fnn$):             FtTmp = FfnTmp(".txt", Fdr, Fnn):                        End Function
Function FxTmp$(Optional Fdr$, Optional Fnn$):             FxTmp = FfnTmp(".xlsx", Fdr, Fnn):                       End Function
Function FxmTmp$(Optional Fdr$, Optional Fnn0$):          FxmTmp = FfnTmp(".xlsm", Fdr, Fnn0):                      End Function
Function FxaTmp$(Optional Fdr$, Optional Fnn0$):          FxaTmp = FfnTmp(".xlsa", Fdr, Fnn0):                      End Function
Function FhtmlTmp$(Html$, Optional PfxFn$)
Dim Ft$: Ft = FfnTmp(".Html", "Html", PfxFn)
End Function
Sub BrwPthTmp():     BrwPth PthTmp:     End Sub
Sub BrwPthTmpRoot(): BrwPth PthTmpRoot: End Sub

Function PthTmpRoot$()
Static X$: If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
PthTmpRoot = X
End Function

Function PthTmp$()
Static X$: If X = "" Then X = PthTmpRoot & "JC\": PthEns X
PthTmp = X
End Function
Function PthTmpFdr$(Fdr):             PthTmpFdr = PthAddFdrEns(PthTmp, Fdr):     End Function
Function PthTmpInst$(Optional Fdr$): PthTmpInst = PthInst(PthEns(PthTmp & Fdr)): End Function
