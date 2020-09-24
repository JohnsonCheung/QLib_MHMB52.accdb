VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_Skuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit


Function SkuyWrongStm(Skuy$(), Stm$) As String(): SkuyWrongStm = SyMinus(Skuy, SkuyStm(Stm)): End Function

Function SkuyStm(Stm$) As String(): SkuyStm = DcStrQC(WSql(Stm)): End Function
Function SkuyMH() As String():       SkuyMH = SkuyStm("M"):       End Function
Function SkuyUD() As String():       SkuyUD = SkuyStm("U"):       End Function

Private Function WSql$(Stm$):    WSql = WSqlOp(IIf(Stm = "M", "<>", "=")): End Function
Private Function WSqlOp$(Op$): WSqlOp = RplQ(WSqlTp, Op):                  End Function
Private Function WSqlTp$()
WSqlTp = "Select Sku" & _
" from Sku" & _
" where Topaz in (" & _
" Select Topaz" & _
" From Topaz" & _
" where Left(CdTopaz,3) ? 'UDV')"
End Function

Function SkuyTbSku() As String(): SkuyTbSku = DcStrTF(CDb, "Sku.Sku"): End Function
