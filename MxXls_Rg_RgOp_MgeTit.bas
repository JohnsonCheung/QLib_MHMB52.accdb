Attribute VB_Name = "MxXls_Rg_RgOp_MgeTit"
Option Compare Text
Const CMod$ = "MxXls_Rg_Op_MgeTit."
Option Explicit

Sub MgeTit(Tit As Range)
Minvn Tit.Application '<-- assume it has Maxv, otherwise, it will break
WsRg(Tit).Activate
Tit.UnMge
Dim R%: For R = 1 To Tit.Rows.Count
    WMgeRow_2 RgR(Tit, R)
Next
Tit.BorderAround XlLineStyle.xlContinuous, xlThick
With Tit.Borders(xlInsideHorizontal)
    .LineStyle = XlLineStyle.xlContinuous
    .Weight = XlBorderWeight.xlThick
End With
With Tit.Borders(xlInsideVertical)
    .LineStyle = XlLineStyle.xlContinuous
    .Weight = XlBorderWeight.xlThick
End With
End Sub
Private Sub WMgeRow_2(Row As Range)
Dim I: For Each I In Itr(W2_RgyMgable_3(Row))
    With CvRg(I)
        .Mge
        .HorizontalAlignment = XlHAlign.xlHAlignCenter
    End With
Next
End Sub
Private Function W2_RgyMgable_3(Row As Range) As Range()
Dim Dr(): Dr = DrFstRg(Row)
Dim Ay() As P12: Ay = W3_P12yMgable_4(Dr)
Dim RowC%: RowC = Row.Column
Dim Ws As Worksheet: Set Ws = Row.Parent
Dim R&: R = Row.Row
Dim J%: For J = 0 To UbP12(Ay)
    Dim M As P12: M = Ay(J)
    Dim C1%: C1 = M.P1 + RowC
    Dim C2%: C2 = M.P2 + RowC
    PushObj W2_RgyMgable_3, RgWsRCC(Ws, R, C1, C2)
Next
End Function
Private Function W3_P12yMgable_4(Dr()) As P12()
Dim C%: For C = 0 To UB(Dr) - 1 ' Mgable C must have next, so -1
    If W4_IsMgable(Dr, C) Then
        PushP12 W3_P12yMgable_4, P12(C, W4_MgableC2(Dr, C))
    End If
Next
End Function
Private Function W4_IsMgable(Dr(), C) As Boolean
If IsEmpty(Dr(C)) Then Exit Function
W4_IsMgable = IsEmpty(Dr(C + 1))
End Function
Private Function W4_MgableC2%(Dr(), C)
Dim U&: U = UB(Dr)
Dim O%: For O = C + 1 To U
    If Not IsEmpty(Dr(O)) Then W4_MgableC2 = O - 1: Exit Function
Next
W4_MgableC2 = U
End Function
