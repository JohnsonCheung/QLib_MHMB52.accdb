Attribute VB_Name = "MxIde_Dcl_Ud"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Ud."
Enum eOptCpr: eOptCprNone: eOptCprTxt: eOptCprBin: eOptCprDb: End Enum
Type VbCnst: Cnstn As String: IsPrv As Boolean: Tyc As String: Tyn As String: V As String: End Type
Type VbVar: Varn As String: IsPrv As Boolean: IsAy As Boolean: Tyc As String: Tyn As String: End Type
Type VbTEmbr: Mbn As String: Enmv As Long: End Type
Type VbEnm: EnmnLn As String: IsPrv As Boolean: Mbr() As VbTEmbr: End Type
Type VbDcl
    OptExp As Boolean
    OptCpr As eOptCpr
    OptBas As Byte
    CnsT() As VbCnst
    Var() As VbVar
    Udt() As TUdt
    Enm() As VbEnm
End Type
Type VbMth
    Mth As Msig
End Type
Type VbMd
    Mdn As String
    Dcl As VbDcl
    Mth() As VbMth
    LasRmk() As String
End Type
Type VbPj
    Pjn As String
    Pjf As String
    Md() As VbMd
End Type
