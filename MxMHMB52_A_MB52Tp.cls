VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_MB52Tp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_MB52Tp."
Private Const W1Msg$ = "After edit and save, use MHMB52Io.EnsTp to import the Tp in PthTp back to Tb-Att"
Function Pth$():      Pth = PthTp:                                        End Function
Function Fn$():        Fn = "TpMB52.xlsx":                                End Function
Function Tp$():        Tp = Pth & Fn:                                     End Function
Sub Opn():                  MaxvFx Tp:                                    End Sub
Sub Edt():                  W1Pmpt: EdtAttTpC "MHMB52Io.Tp.xlsx":         End Sub ' edit the template in Tb-Att. (1) Exp to PthTp (2) opn.  Then, user needs to MHMB52Io.EnsTp
Private Sub W1Pmpt():       MsgBox W1Msg, vbInformation:                  End Sub
Sub Fmt():                  Static X As New MxMHMB52_A_MB52Tp_Fmt: X.Fmt: End Sub
Sub ImpIf():                ImpAttTpC Tp, Fn:                             End Sub
Sub Exp(FfnTo$):            ExpAttTpC Fn, FfnTo:                          End Sub
