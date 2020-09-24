Attribute VB_Name = "MxXls_Wc_Op_Add"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Wc_Op_Add."

Private Sub B_CrtFxWc()
Dim Fx$: Fx = FxTmp
CrtFxWc Fx, CFb, TnyOupCn(CurrentProject.Connection)
MaxvFx Fx
End Sub
Sub CrtFxWc(Fx$, Fb$, Tny$())
Const CSub$ = CMod & "CrtFxWc"
ChkFfnNExi Fx, CSub, "Excel file"
X_1WbNw_2(Fb, Tny).SaveAs Fx
End Sub

Private Sub B_CrtFxFbOup()
Dim Fx$: Fx = FxTmp
CrtFxFbOup Fx, CFb
MaxvFx Fx
End Sub
Sub CrtFxFbOupC(Fx$): CrtFxFbOup Fx, CFb: End Sub
Sub CrtFxFbOup(Fx$, Fb$)
Dim B As Workbook: Set B = X_1WbNw_2(Fb, TnyOup(Db(Fb)))
W_2CrtWsByAllWc B
Dim S As Worksheet: Set S = WsFst(B)
S.Name = "Index"
Stop 'EnsWsHyp A1Ws(S)
B.SaveAs Fx
End Sub
Private Sub W_2CrtWsByAllWc(B As Workbook)
Dim C As WorkbookConnection: For Each C In B.Connections
    WCrtWs B, C
Next
End Sub
Private Sub WCrtWs(B As Workbook, Wc As WorkbookConnection)
'Assume @Wc.Name is @xxxx so that Lo.Name is using Lo_{RmvFst(Wc.Name)}
Dim S As Worksheet: Set S = B.Sheets.Add
Dim N$: N = Wc.Name
Dim Cns$: Cns = Wc.OLEDBConnection.Connection
S.Range("A1").Value = "Back"
With S.ListObjects.Add(SourceType:=xlSrcExternal, Source:=Cns, Destination:=S.Range("A2")).QueryTable
    .CommandType = xlCmdTable
    .CommandText = N
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnInfo = True
    .ListObject.DisplayName = "Lo_" & RmvFst(N)
    .Refresh BackgroundQuery:=False
End With
S.Name = N
End Sub

Private Function X_1WbNw_2(Fb$, Tny$()) As Workbook
Dim B As Workbook: Set B = WbNw
X2_AddWcTny_3 B, Fb, Tny
Set X_1WbNw_2 = B
End Function
Private Sub X2_AddWcTny_3(B As Workbook, Fb$, Tny$())
Dim T: For Each T In Tny
    X3_AddWc B, Fb, T
Next
End Sub
Private Sub X3_AddWc(B As Workbook, Fb$, T)
B.Connections.Add2 T, T, CnsFbOle(Fb), T, XlCmdType.xlCmdTable
End Sub
