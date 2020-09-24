Attribute VB_Name = "MxIde_Dta_ePPrv"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dta_eWhPPrv."
Enum eWhPPrv: eWhPPrvPub: eWhPPrvPrv: eWhPPrvBoth: End Enum
Public Const EnmqssPPrv$ = "eWhPPrv? Prv Pub Both"
Function HitPPrv(IsPrv As Boolean, PuPrv As eWhPPrv) As Boolean
Select Case PuPrv
Case eWhPPrvBoth: HitPPrv = True
Case eWhPPrvPrv: HitPPrv = IsPrv
Case eWhPPrvPub: HitPPrv = Not IsPrv
Case Else: ThwEnm CSub, PuPrv, EnmqssPPrv
End Select
End Function
Function eWhPPrvIsPPrv2(IsPub As Boolean, IsPrv As Boolean) As eWhPPrv
Dim O As eWhPPrv
Select Case True
Case (IsPub And IsPrv) Or (Not IsPub And Not IsPrv): O = eWhPPrvBoth
Case IsPub: O = eWhPPrvPub
Case IsPrv: O = eWhPPrvPrv
End Select
eWhPPrvIsPPrv2 = O
End Function
