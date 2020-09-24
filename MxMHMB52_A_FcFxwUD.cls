VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcFxwUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcFxwUD."
Friend Sub A_OpnFxiLas():   MaxvFx MH.FcLas.FxiUDPthi: End Sub
Friend Sub B_WCC_ErM0115(): WCC_ErM0115__Tst:          End Sub
Friend Sub B_CrtFxw()
Dim A As StmYm: A = MH.FcLas.StmYmUDPthi
CrtFxw A
Dim Fxi$
Dim Fxw$
    Fxi = MH.FcIO.Fxi(A)
    Fxw = MH.FcIO.Fxw(A)
MH.FcLas.StmYmUDPthi
OpnFxAp Fxi, Fxw
End Sub
Sub CrtFxw(A As StmYm)
'Inp: MH.FcIO.Fxi(A) for UD.  All Sku be belong UD
'Oup: MH.FcIO.Fxw(A)(Same stru as FcSku: VerYY,VerMM,YY,MM,Co,Sku,SC.  Verify all SKU should be UD, anything wrong report and raise error
Dim Fxi$
    Fxi = MH.FcIO.Fxi(A)
    ChkFxww Fxi, MH.FcIO.Wsnn("U")
Dim Fxw$
    Fxw = MH.FcIO.Fxw(A)
    DltFfnIf Fxw
    FileCopy Fxi, Fxw
W_ChkCell_M0115__CC Fxi, A
Dim B As Workbook: Set B = WbFx(Fxw)
    W_FmtWbFxw__WFW WWWs86(B)
    W_FmtWbFxw__WFW WWWs87(B)
    B.Close True
End Sub
Private Sub W_FmtWbFxw__WFW(S As Worksheet)
WFW_Dlt_row_and_col_aft S 'Delete Row and DcDrs
WFW_Put_MMM_YYYY S             'Put Row 1 COl-B 15 columns to Row 2==
WFW_RenCell_As_M0115 S             'Ren Field as : Sku M01..M15
End Sub
Private Sub WFW_Dlt_row_and_col_aft(S As Worksheet)
S.Range("$A:$C").EntireColumn.Delete
S.Range("$B:$AC").EntireColumn.Delete
DltColFm S, "P"
End Sub
Private Sub WFW_Put_MMM_YYYY(S As Worksheet)
Dim Rg As Range, V
Dim J%: For J = 1 To 15
    Set Rg = S.Cells(1, J + 1)
    V = Rg.Value
    Set Rg = S.Cells(2, J + 1)
    Rg.Value = V
Next
S.Range("$1:$1").EntireRow.Delete
S.Range("$1:$1").NumberFormat = "MMM YYYY"
End Sub
Private Sub W_ChkCell_M0115__CC(Fxi$, A As StmYm)
Dim B As Workbook: Set B = WbFx(Fxi)
Dim E1$(): E1 = WCC_ErM0115(WWWs86(B), A)
Dim E2$(): E2 = WCC_ErM0115(WWWs87(B), A)
Dim E$(): E = SyAddAp(E1, E2)
If Si(E) = 0 Then B.Close False: Exit Sub
MaxvWb B
ChkEry E
End Sub
Private Function WWWs86(B As Workbook) As Worksheet:             Set WWWs86 = WWWsCo(B, 86):               End Function
Private Function WWWs87(B As Workbook) As Worksheet:             Set WWWs87 = WWWsCo(B, 87):               End Function
Private Function WWWsCo(B As Workbook, Co As Byte) As Worksheet: Set WWWsCo = B.Sheets(MH.FcIO.WsnUD(Co)): End Function
Private Sub WCC_ErM0115__Tst()
Dim A As StmYm: A = MH.FcLas.StmYmMHPthi
Dim Fxi$
    Fxi = MH.FcIO.Fxi(A)
Dim B As Workbook: Set B = WbFx(Fxi)
Dim E1$(), E2$()
    E1 = WCC_ErM0115(WWWs86(B), A)
    E2 = WCC_ErM0115(WWWs87(B), A)
ChkEry SyAdd(E1, E2)
ClsWbNoSav B
End Sub
Private Function WCC_ErM0115(S As Worksheet, A As StmYm) As String()
Dim O$(), R As Range
    Dim Dtey() As Date: Dtey = DteyYm(Ym(A.Y, A.M))
    Dim DrAGAU(): DrAGAU = DrSq(S.Range("AG1:AU1").Value)
    Dim RgAG As Range: Set RgAG = S.Range("AG1")
    Dim J%: For J = 0 To 14
        PushNB O, WCC_ErIx(Dtey(J), DrAGAU(J), J, RgAG)
    Next
If Si(O) = 0 Then Exit Function
Dim OO$()
Dim Wsn$: Wsn = S.Name
Dim B As Workbook: Set B = WbWs(S)
PushS OO, "Following error(s) are in the forecast import file, fix them and import again:"
PushS OO, "Folder    : [" & Pth(B.FullName) & "]"
PushS OO, "Excel File: [" & B.Name & "]"
PushS OO, "Worksheet : [" & Wsn & "]"
PushS OO, ""
PushIAy OO, O
PushS OO, ""
WCC_ErM0115 = OO
End Function
Private Function WCC_ErIx$(DteEpt As Date, ValCell, Ix%, RgAG As Range)
Dim Adr$
If Not IsDte(ValCell) Then
    Adr = RgRC(RgAG, 1, Ix + 1).Address
    WCC_ErIx = FmtQQ("Cell[?] should be a [Date], but now it is a [?]", Adr, TypeName(ValCell))
    Exit Function
End If
If ValCell <> DteEpt Then
    Adr = RgRC(RgAG, 1, Ix + 1).Address
    WCC_ErIx = FmtQQ("Cell[?] should be a [Date] of [?], but now it is [?]", Adr, DteEpt, ValCell)
End If
End Function

Private Sub WFW_RenCell_As_M0115(Ws As Worksheet)
Ws.Range("A1").Value = "Sku"
Dim J%, Rg As Range: For J = 1 To 15
    Set Rg = Ws.Cells(1, 1 + J)
    Rg.Value = "M" & Format(J, "00")
Next
End Sub
