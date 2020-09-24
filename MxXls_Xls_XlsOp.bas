Attribute VB_Name = "MxXls_Xls_XlsOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Op_."
Enum ePosWs: ePosWsEnd: ePosWsBeg: ePosWsBef: ePosWsAft: End Enum
Sub ClsAllXls()
Dim X As Excel.Application
Do
    Set X = XlsGet
    If IsNothing(X) Then Exit Sub
    X.Visible = True
    X.WindowState = xlMaximized
    X.Quit
Loop
End Sub
Sub BrwFx(Fx)
MaxvFx Fx
End Sub

Sub CrtFx(Fx$): WbSavAs(WbNw, Fx).Close: End Sub

Function WbEns(Fx$) As Workbook
If HasFfn(Fx) Then
    Set WbEns = WbFx(Fx)
Else
    Set WbEns = WbNw
    WbEns.SaveAs Fx
End If
End Function
Sub EnsFx(Fx$)
If NoFfn(Fx) Then CrtFx Fx
End Sub

Function RnyOfEmpRowFmSq(RCnoSq()) As Long()
Dim LCno%: LCno = LBound(RCnoSq, 2)
Dim UCno%: UCno = UBound(RCnoSq, 2)
Dim IsEmpRow As Boolean
Dim Rno&: For Rno = LBound(RCnoSq, 1) To UBound(RCnoSq, 1)
    GoSub IsEmpRow
    If IsEmpRow Then PushI RnyOfEmpRowFmSq, Rno
Next
Exit Function
IsEmpRow: '(RCnoSq,Rno,UCno,LCno)
    Dim Cno%: For Cno = LCno To UCno
        If Not IsEmpty(RCnoSq(Rno, Cno)) Then IsEmpRow = False: Return
    Next
    IsEmpRow = True
    Return
End Function

Function AdrRny$(RnyOrdered)
Select Case Si(RnyOrdered)
Case 0: Thw CSub, "RnyOrdered does not have element"
Case 1: AdrRny = RplQ("$?:$?", RnyOrdered(0)): Exit Function
End Select
Dim O$()
    Dim L12y() As L12: L12y = L12yNumy(RnyOrdered)
    Dim J&: For J = 0 To UbL12(L12y)
        PushI O, FmtQQ("$?:$?")
    Next
AdrRny = JnSemi(O)
End Function
Function EntRows(S As Worksheet, RnyOrdered) As Range: Set EntRows = S.Range(AdrRny(RnyOrdered)): End Function
Sub RgRmv_NRow(At As Range, Optional N = 1):                         RgRREnt(At, 1, N).Delete:    End Sub
Private Sub B_SetWscdn()
Dim S As Worksheet: Set S = WsNw
SetWsCdn S, "XX"
Maxv S.Application
End Sub

Sub MgeBottomCell(BarC As Range)
Ass IsRgSngCol(BarC)
Dim R2: R2 = NRowRg(BarC)
Dim R1
    Dim Fnd As Boolean
    For R1 = R2 To 1 Step -1
        If Not IsEmpty(RgRC(BarC, R1, 1)) Then Fnd = True: GoTo Nxt
    Next
Nxt:
    If Not Fnd Then Stop
If R2 = R1 Then Exit Sub
Dim R As Range: Set R = RgCRR(BarC, 1, R1, R2)
R.Merge
R.VerticalAliment = XlVAlign.xlVAlignTop
End Sub

Sub SetWsCdn(S As Worksheet, CdNm$): CmpWs(S).Name = CdNm: End Sub

Sub SetWsCdnAndLon(S As Worksheet, Nm$)
CmpWs(S).Name = Nm
SetLon LoFst(S), Nm
End Sub

Function HasCell(R As Range, Cell As Range) As Boolean
If Not IsCell(Cell) Then Exit Function
If Not IsBet(Cell.Row, R.Row, NRowRg(R)) Then Exit Function
If Not IsBet(Cell.Column, R.Column, NColRg(R)) Then Exit Function
HasCell = True
End Function

Sub UnMgeRg(R As Range)
Minvn R.Application
R.UnMerge
End Sub
Sub MgeRg(R As Range)
Minv R.Application
R.Merge ' same as R.MergeCells = True
R.HorizontalAliment = XlHAlign.xlHAlignCenter
R.VerticalAliment = Excel.XlVAlign.xlVAlignCenter
End Sub

Sub ClsWbAllNoSav()
Dim X As Excel.Application: Set X = Xls
While X.Workbooks.Count > 0
    ClsWbNoSav X.Workbooks(1)
Wend
End Sub

Sub ClsWbNoSav(B As Workbook):  B.Close False:         End Sub
Sub ClsWsNoSav(S As Worksheet): WbWs(S).Close False:   End Sub
Sub QuitWb(B As Workbook):      QuitXls B.Application: End Sub
Sub QuitWs(S As Worksheet):     QuitXls S.Application: End Sub

Function SavWbCsv(A As Workbook, Fcsv$) As Workbook
DltFfnIf Fcsv
A.Application.DisplayAlerts = False
A.SaveAs Fcsv, XlFileFormat.xlCSV
A.Application.DisplayAlerts = True
Set SavWbCsv = A
End Function

Sub SavWb(B As Workbook)
Dim Y As Boolean
Y = B.Application.DisplayAlerts
B.Application.DisplayAlerts = False
B.Save
B.Application.DisplayAlerts = Y
End Sub

Function WbSavAs(A As Workbook, Fx, Optional Fmt As XlFileFormat = xlOpenXMLWorkbook) As Workbook
Dim Y As Boolean
Y = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
A.SaveAs Fx, Fmt
A.Application.DisplayAlerts = Y
Set WbSavAs = A
End Function

Sub SetWcFcsv(B As Workbook, Fcsv$)
'Set first B TextConnection to Fcsv if any
Dim T() As TextConnection: T = WcyTxtzWb(B)
Stop '
'Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
'T.Connection = "TEXT;" & Fcsv
End Sub

Function HasWs(A As Workbook, Wsix) As Boolean
If IsNumeric(Wsix) Then
    HasWs = IsBet(Wsix, 1, A.Sheets.Count)
    Exit Function
End If
Dim S As Worksheet
For Each S In A.Worksheets
    If S.Name = Wsix Then HasWs = True: Exit Function
Next
End Function

Private Sub B_SetWcFcsv()
Dim B As Workbook
'Set B = WbFx(CVbe_MthFx)
Stop 'Debug.Print TxtWcsy_Wb(B)
SetWcFcsv B, "C:\ABC.CSV"
Stop 'Ass TxtWcsy_Wb(B) = "TEXT;C:\ABC.CSV"
B.Close False
Stop
End Sub

Sub DltLo(S As Worksheet)
While S.ListObjects.Count > 0
    S.ListObjects(0).Delete
Wend
End Sub

Sub DltWsIf(A As Workbook, Wsix)
If HasWs(A, Wsix) Then DltWs A, Wsix
End Sub

Sub SavAsCls(B As Workbook, Fx):  B.SaveAs Fx: B.Close:                                     End Sub
Sub SavAsFxm(B As Workbook, Fxm): B.SaveAs Fxm, XlFileFormat.xlOpenXMLWorkbookMacroEnabled: End Sub

Function SavAsTmpFxm$(B As Workbook)
Dim O$: O = FxmTmp
SavAsFxm B, O
SavAsTmpFxm = O
End Function

Function WbnWs$(S As Worksheet): WbnWs = WbWs(S).FullName: End Function

Sub SetPivFldOri(T As PivotTable, FF$, Ori As XlPivotFieldOrientation)
Dim Tot: Tot = Array(False, False, False, False, False, False, False, False, False, False, False, False)
Dim J%: J = 1
Dim F: For Each F In Itr(FnyFF(FF))
    With PivFld(T, F)
        .Orientation = Ori
        .Position = J
        If Ori = xlColumnField Or Ori = xlRowField Then
            .Subtotals = Tot
        End If
    End With
    J = J + 1
Next
End Sub

Sub ChkHasWsn(B As Workbook, Wsn$, Fun$)
If HasWs(B, Wsn) Then
    Thw Fun, "Wb should have not have Ws", "Wb Ws", B.FullName, Wsn
End If
End Sub

Sub QuitXlsAll()
Dim X As Excel.Application
Do
    Dim J%: ThwLoopTooMuch CSub, J, 100
    Set X = XlsGet: If IsNothing(X) Then Exit Sub
    QuitXls X
Loop
End Sub
Sub QuitXls(X As Excel.Application)
Stamp "QuitXls: Start"
Stamp "QuitXls: ClsWbAll":    ClsWbAll X
Stamp "QuitXls: Quit":        X.Quit
Stamp "QuitXls: Set nothing": Set X = Nothing
Stamp "QuitXls: Done"
End Sub

Sub ClsWbAll(X As Excel.Application)
Dim B As Workbook: For Each B In X.Workbooks
    B.Close False
Next
End Sub

Private Sub B_CrtFxTbo()
Dim Fx$: Fx = FxTmp
Stop 'CrtFxTbOup Fx, FbDtaMHDuty
MaxvFx Fx
End Sub
Sub CrtFxTbOup(Fx, Fb, Optional Way As eLoAddTblWay): SavAsCls WbFbOup(Fb, Way), Fx: End Sub

Sub PutSnoDown(At As Range, N, Optional Fm = 1)
PutAyVert LngySno(N - 1, Fm), At
End Sub

Sub DltSheet1(B As Workbook)
DltWs B, "Sheet1"
End Sub
Sub ActWs(S As Worksheet)
If IsEqObj(S, CWs) Then Exit Sub
S.Activate
End Sub
Sub DltWs(B As Workbook, Wsix)
B.Application.DisplayAlerts = False
If B.Sheets.Count = 1 Then Exit Sub
If HasWs(B, Wsix) Then WsWb(B, Wsix).Delete
End Sub
Sub DltWsnn(B As Workbook, Wsnn$): Stop 'DltWny B, Tml(Wsnn): End Sub

End Sub
Sub DltWny(B As Workbook, Wny$())
Dim N: For Each N In Itr(Wny)
    DltWs B, N
Next
End Sub
Sub ClrDown(R As Range)
RgAtDown(R).Clear
End Sub

Sub FillSnoDown(BarC As Range)
Dim Sq()
Sq = SqrSno(NRowRg(BarC))
RgSq(Sq, BarC).Value = Sq
End Sub
Sub ClrCellBelow(Cell As Range)
CellBelow(Cell).Clear
End Sub

Sub FillWny(At As Range)
RgAyV Wny(WbRg(At)), At
End Sub

Sub FillAtV(At As Range, Ay)
FillSq SqCol(Ay), At
End Sub

Sub FillLc(Lo As ListObject, Coln$, Ay)
Const CSub$ = CMod & "FillLc"
If NRowRg(Lo.DataBodyRange) <> Si(Ay) Then Thw CSub, "Lo-NRow <> Si(Ay)", "Lo-NRow ,Si(Ay)", NRowLo(Lo), Si(Ay)
Dim At As Range, C As ListColumn, R As Range
'DmpAy FnyLo(Lo)
'Stop
Set C = Lo.ListColumns(Coln)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
FillAtV At, Ay
End Sub
Sub FillSq(Sq(), At As Range)
RgSq(Sq, At).Value = Sq
End Sub
Sub FillAtH(Ay, At As Range)
FillSq SqRow(Ay), At
End Sub

Sub RunFxqByCn(Fx, Q)
CnFx(Fx).Execute Q
End Sub
Function DKValKSet(KSet As Dictionary) As Drs
Dim K, Dy(): For Each K In KSet.Keys
    Dim Sset As Dictionary: Set Sset = KSet(K)
    Dim V: For Each V In Sset.Keys
        PushI Dy, Array(K, V)
    Next
Next
DKValKSet = DrsFf("K V", Dy)
End Function
Private Sub B_DKValLoFilter()
Dim Lo As ListObject: Set Lo = LoFst(CWs)
BrwDrs DKValLoFilter(Lo)
End Sub
Function DKValLoFilter(L As ListObject) As Drs
DKValLoFilter = DKValKSet(AetKeyLoFilter(L))
End Function

Function AetKeyKyAetAy(Ky$(), AetAy() As Dictionary) As Dictionary
Set AetKeyKyAetAy = New Dictionary
Dim K, J&: For Each K In Itr(Ky)
    AetKeyKyAetAy.Add K, AetAy(J)
    J = J + 1
Next
End Function
Sub SetOnFilter(L As ListObject)
On Error GoTo X
Dim M As Boolean: M = L.AutoFilter.FilterMode ' If filter is on, it will have no err, otherwise, there is err
Exit Sub
X:
L.Range.AutoFilter 'Turn on
End Sub
Function AetKeyLoFilter(L As ListObject) As Dictionary
'Ret : KSet
Dim O As Dictionary: Set O = New Dictionary
SetOnFilter L
Dim Fny$(): Fny = FnyLo(L)
Dim F As Filter, J%: For Each F In L.AutoFilter.Filters
    Dim K$: K = Fny(J)
    AetKeyLoFilter__Add O, K, F
    J = J + 1
Next
Set AetKeyLoFilter = O
End Function

Sub AetKeyLoFilter__Add(OKSet As Dictionary, K$, F As Filter)
If Not F.On Then Exit Sub
If F.Operator <> xlFilterValues Then Exit Sub
Dim S As Dictionary: Set S = AetAy(AmRmvPfx(F.Criteria1, "="))
OKSet.Add K, S
End Sub
