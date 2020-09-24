Attribute VB_Name = "MxVb_Ay_Op_FmtCntAy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_OpCnt."
Enum eCntDup: eCntDupAll: eCntDupOnly: eCntDupSng: End Enum
Enum eCntSrtOpt: eNoSrt: eSrtByCnt: eSrtByItm: End Enum

Function FmtCntAy(Ay, Optional Opt As eCntDup, Optional SrtOpt As eCntSrtOpt, Optional By As eSrt) As String()
Const CSub$ = CMod & "FmtCntAy"
Dim D As Dictionary: Set D = DiCnt(Ay, Opt)
Dim K
Dim W%: W = DicItmWdt(D)
Dim O$()
Select Case SrtOpt
Case eNoSrt
    FmtCntAy = CntLyCntDi(D, W)
Case eSrtByCnt
    FmtCntAy = AySrtQ(CntLyCntDi(D, W), By)
Case eSrtByItm
    FmtCntAy = CntLyCntDi(DiSrt(D, By), W)
Case Else
    Thw CSub, "Invalid SrtOpt", "SrtOpt", SrtOpt
End Select
End Function

Sub BrwCntAy(Ay, Optional Opt As eCntDup): Brw FmtDi(DiCnt(Ay, Opt)): End Sub
