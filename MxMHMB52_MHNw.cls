VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_MHNw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_MH__Nw."
Private Fun As New MxMHMB52_MHFun
Function StmYm(Stm$, Y As Byte, M As Byte) As StmYm
Fun.ChkStm Stm
Select Case True
Case Y > 99:         ThwPm "StmYm", "Y", Y, "0-99"
End Select
With StmYm
    .Stm = Stm
    .Y = Y
    .M = M
End With
End Function
Function StmYmzYm(Stm$, A As Ym) As StmYm
With StmYmzYm
    .Stm = Stm
    .Y = A.Y
    .M = A.M
End With
End Function
Function CoStmYm(Co As Byte, Stm$, Y As Byte, M As Byte) As CoStmYm
With CoStmYm
    .Co = Co
    .Stm = Stm
    .M = M
    .Y = Y
End With
End Function
Function CoYmd(Co As Byte, A As Ymd) As CoYmd
Fun.ChkCo Co
CoYmd.Co = Co
CoYmd.Ymd = A
End Function
Function CoYmdzLas(Co As Byte) As CoYmd: CoYmdzLas = CoYmd(Co, MH.TbOH.YmdLas): End Function
Function CoYm(Co As Byte, Y As Byte, M As Byte) As CoYm
Fun.ChkCo Co
With CoYm
    .Co = Co
    .Y = Y
    .M = M
End With
End Function
