VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcFxwMH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcFxwMH."
Function TyFc$(): TyFc = "Final Forecast M-1": End Function
Function Wsn$():   Wsn = "Table":              End Function
Friend Sub B_CrtFxw():
Dim M  As StmYm: M = MH.FcLas.StmYmMH
CrtFxw M
OpnFxAp MH.FcIO.Fxw(M), MH.FcIO.Fxi(M)
End Sub
Sub CrtFxwLas(): CrtFxw MH.FcLas.StmYmMH: End Sub
Sub CrtFxw(A As StmYm)
Dim Fxi$, Fxw$
    Fxi = MH.FcIO.Fxi(A)
    Fxw = MH.FcIO.Fxw(A)
'== Kill & Copy
DltFfnIf Fxw
FileCopy Fxi, Fxw

'== Open
Dim B As Workbook: Set B = WbFx(Fxw)
Stop
'== Delet Sheet
B.Application.DisplayAlerts = False
Dim J%, S As Worksheet
For J = B.Sheets.Count To 1 Step -1
    Set S = B.Sheets(J)
    If S.Name <> MH.FcIO.WsnMH Then S.Visible = xlSheetVisible: S.Delete
Next
'== Delete Row and DcDrs
Set S = B.Sheets(MH.FcIO.WsnMH)
S.Range("$1:$14").EntireRow.Delete
S.Range("$A:$F").EntireColumn.Delete
S.Range("$C:$E").EntireColumn.Delete
S.Range("$D:$D").EntireColumn.Delete
S.Range("$E:$E").EntireColumn.Delete

'== Delete Shapes
While S.Shapes.Count > 0
    S.Shapes(1).Delete
Wend
'== Column D == FcTyMH
S.Range("D1").Value = "FcTyMH"
'== Insert Listobject
Dim C2R&: C2R = S.Cells.SpecialCells(xlCellTypeLastCell).Row
Dim C1 As Range: Set C1 = S.Range("A1")
Dim C2 As Range: Set C2 = S.Cells(C2R, "P")
Dim Rg As Range: Set Rg = S.Range(C1, C2)
Dim Lo As ListObject: Set Lo = S.ListObjects.Add(xlSrcRange, Rg, , xlYes)
Lo.TableStyle = "TableStyleLight1"
Lo.Range.AutoFilter Field:=4, Criteria1:=TyFc

'== ErCellVal
Dim O$()
    Dim M As Ym: M = Ym(A.Y, A.M)
    Dim NyM$(): NyM = NyMonthM3SpcY4(M)
    PushNB O, ErCellVal(S, "A1", "Market")
    PushNB O, ErCellVal(S, "B1", "Market Channel")
    PushNB O, ErCellVal(S, "C1", "Product")
    PushNB O, ErCellVal(S, "D1", "FcTyMH")
    
    For J = 0 To 11
        Dim Adr$: Adr = RgWsRC(S, 1, 5 + J).Address(False, False)
        PushNB O, ErCellVal(S, Adr, NyM(J))
    Next
    
    If Si(O) Then
        Dim OO$()
        PushS OO, "Following error(s) are in the forecast import file, fix them and import again:"
        PushS OO, "Folder    : [" & Pth(B.FullName) & "]"
        PushS OO, "Excel File: [" & B.Name & "]"
        PushS OO, "Worksheet : [" & Wsn & "]"
        PushS OO, ""
        BrwAy SyAdd(OO, O)
        B.Save
        Maxv B.Application
        RaiseNotePad
    End If

'== RenSku
'== RenCoNm
Set Rg = S.Range("C1"): Rg.Value = "Sku"
Set Rg = S.Range("B1"): Rg.Value = "CoNm"
'== Ren M01..12
For J = 1 To 12
    Set Rg = S.Cells(1, 4 + J)
    Rg.Value = "M" & Format(J, "00")
Next

'== Save / Close / Quit Xls
SavWbQuit B
End Sub
