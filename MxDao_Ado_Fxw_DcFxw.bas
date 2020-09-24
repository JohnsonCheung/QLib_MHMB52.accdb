Attribute VB_Name = "MxDao_Ado_Fxw_DcFxw"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Fxw_DcFxw."

Function DcStrFxq(Fx$, Q$) As String(): DcStrFxq = DcStrArs(ArsFxq(Fx, Q)): End Function
Function DcStrDisFxw(Fx$, W$, C$) As String()
Dim Q$: Q = SqlSelFld(C, Axtn(W))
DcStrDisFxw = DcStrArs(ArsFxq(Fx, Q))
End Function
Function DcStrFxw(Fx$, W$, Coln$, Optional Bepr$) As String(): DcStrFxw = intoDcFxw(SyEmp, Fx, W, Coln, Bepr): End Function

Function DcFxq(Fx$, Q$) As Variant(): DcFxq = DcArs(ArsFxq(Fx, Q)): End Function


Function DcIntFxw(Fx$, W$, C$) As Integer():                        DcIntFxw = intoDcFxw(IntyEmp, Fx, W, C):                 End Function
Function DcFxw(Fx$, W$, C$) As Variant():                              DcFxw = intoDcFxw(AvEmp, Fx, W, C):                   End Function ' :Av '#DcDrs-Value-Ay#
Private Function intoDcFxw(Intoy, Fx$, W$, Coln$, Optional Bepr$): intoDcFxw = DcIntoArs(Intoy, ArsFxwc(Fx, W, Coln, Bepr)): End Function

Private Sub HasFxwcBlnk__Tst()
GoSub T1
Exit Sub
Dim Fx$, W$, ColnStr$
T1:
    Fx = sampFfn("SampMB52Fxi.xlsx")
    W = "Sheet1"
    ColnStr = "Material"
    Ept = False
    GoTo Tst
Tst:
    Act = HasFxwcBlnk(Fx, W, ColnStr)
    C
    Return
End Sub
Function HasFxwcBlnk(Fx$, W$, ColnStr$) As Boolean
Dim Q$: Q = FmtQQ("Select Top 1 [?] from [?] where Nz(Trim([?]),'')=''", ColnStr, Axtn(W), ColnStr)
Dim Dc$(): Dc = DcStrFxq(Fx, Q)
HasFxwcBlnk = Si(Dc) = 1
End Function
