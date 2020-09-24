Attribute VB_Name = "MxDao_Att_AttTp"
Option Compare Text
Const CMod$ = "MxDao_Att_AttTp."
Option Explicit
Sub DltAttTp(D As Database, Attfn$): DltAtt D, "Tp", Attfn: End Sub
Sub DltAttTpC(Attfn$):               DltAttTp CDb, Attfn:   End Sub
Sub EdtAttTp(D As Database, Attfn$)
Dim Tp$: Tp = PthTp & Attfn
DltFfnIf Tp
ExpAttTp D, Attfn, Tp
MaxvFx Tp
End Sub
Sub EdtAttTpC(Attfn$):                       EdtAttTp CDb, Attfn$:                   End Sub
Sub ExpAttTp(D As Database, Attfn$, FfnTo$): ExpAtt D, "Tp", Attfn, FfnTo:           End Sub
Sub ExpAttTpC(Attfn$, FfnTo$):               ExpAttTp CDb, Attfn, FfnTo:             End Sub
Sub ExpAttTpIf(D As Database, Attfn$):       ExpAttIf D, PthTp & Attfn, "Tp", Attfn: End Sub
Sub ExpAttTpIfC(Attfn$):                     ExpAttTpIf CDb, Attfn:                  End Sub
Sub ImpAttTp(D As Database, FfnFm$, Attfn$): ImpAtt FfnFm, D, "Tp", Attfn:           End Sub
Sub ImpAttTpC(FfnFm$, Attfn$):               ImpAttTp CDb, FfnFm, Attfn:             End Sub
Sub ImpAttTpIf(D As Database, Attfn$):       ImpAttIf D, PthTp & Attfn, "Tp", Attfn: End Sub
Sub ImpAttTpIfC(Attfn$):                     ImpAttTpIf CDb, Attfn:                  End Sub
