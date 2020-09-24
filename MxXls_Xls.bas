Attribute VB_Name = "MxXls_Xls"
':LoZZ$ = "z when used in Nm, it has special meaning.  It can occur in Cml las-one, las-snd, las-thrid chr, else it is er."
':NmBrk$ = "NmBrk is z or zx or zxx where z is letter-z and x is lowcase or digit.  NmBrk must be sfx of a cml."
':NmBrk_za$ = "It means `and`."
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Xls."
Public Const MaxCno% = 16384
Public Const MaxRno& = 1048576
Public Const FexeXls$ = "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"

Function TRcA1Rg(R As Range) As TRc
With TRcA1Rg
    .R = R.Row
    .C = R.Column
End With
End Function

Function WbWs(A As Worksheet) As Workbook:     Set WbWs = A.Parent:           End Function
Function WsMain(A As Workbook) As Worksheet: Set WsMain = WsCdn(A, "WsMain"): End Function


Private Sub B_XlsGet(): Debug.Print ObjPtr(XlsGet): End Sub
Function XlsGet() As Excel.Application
On Error Resume Next
Set XlsGet = GetObject(, "Excel.Application")
End Function

Function WbFst() As Workbook:                         Set WbFst = WbFstX(Xls):    End Function
Function WbFstX(X As Excel.Application) As Workbook: Set WbFstX = X.Workbooks(1): End Function

Function XlsDft(X As Excel.Application) As Excel.Application
If IsNothing(X) Then
    Set XlsDft = Xls
Else
    Set XlsDft = X
End If
End Function

Function NLc%(L As ListObject):       NLc = L.ListColumns.Count:     End Function
Function NDcLo%(L As ListObject):   NDcLo = L.ListColumns.Count:     End Function
Function NRowLo&(L As ListObject): NRowLo = NRowRg(L.DataBodyRange): End Function

Function RgLoAllCo(L As ListObject) As Range:         Set RgLoAllCo = RgLoCC(L, 1, NDcLo(L)):    End Function
Function RgLoAllColEnt(L As ListObject) As Range: Set RgLoAllColEnt = RgLoAllCo(L).EntireColumn: End Function
Function LonyWb(B As Workbook) As String()
Dim S As Worksheet: For Each S In B.Sheets
    Dim L As ListObject: For Each L In S.ListObjects
        PushI LonyWb, L.Name
    Next
Next
End Function

Function DrsLoFny(L As ListObject, Fny$()) As Drs
DrsLoFny = Drs(Fny, DySqCnoy(SqLo(L), CnoyLoFny(L, Fny)))
End Function

Function DrsLo(L As ListObject) As Drs: DrsLo = Drs(FnyLo(L), DyLo(L)): End Function

Function DyRgCnoy(Rg As Range, Cnoy%()) As Variant():       DyRgCnoy = DySqCnoy(SqRg(Rg), Cnoy):                       End Function
Function DyLoCCC(L As ListObject, CCC$) As Variant():        DyLoCCC = DyRgCnoy(L.DataBodyRange, Cnoy(FnyLo(L), CCC)): End Function ' Return as many column as columns in [CC] from L
Function AdrLoDta$(L As ListObject):                        AdrLoDta = AdrWs(L.DataBodyRange):                         End Function
Function RgEntColLC(L As ListObject, C) As Range:     Set RgEntColLC = RgLc(L, C).EntireColumn:                        End Function ' entire col range

Function LoWsDta(B As Worksheet) As ListObject: Set LoWsDta = LoRg(DtaRg(B)): End Function

Function StrFbtLo$(L As ListObject):         StrFbtLo = FmtFbtQt(L.QueryTable): End Function
Function FnyLo(L As ListObject) As String():    FnyLo = Itn(L.ListColumns):     End Function

Function FfLo$(L As ListObject):                          Stop '   FfLo = TmlAy(FnyLo(L.ListColumns)):             End Function

End Function
Function HasLc(L As ListObject, C$) As Boolean:             HasLc = HasItn(L.ListColumns, C):              End Function
Function IsLoNoDta(L As ListObject) As Boolean:         IsLoNoDta = IsNothing(L.DataBodyRange):            End Function
Function CellLoHdr(L As ListObject, Fldn) As Range: Set CellLoHdr = RgRC(L.ListColumns(Fldn).Range, 1, 1): End Function

Function LcLoC(L As ListObject, C) As ListColumn
Const CSub$ = CMod & "LcLoC"
On Error GoTo X
Set LcLoC = L.ListColumns(C)
Exit Function
X: Thw CSub, "Given-@C is not found @L", "Given-@C Lon FnyLo", C, L.Name, FnyLo(L)
End Function

Function LcFst(L As ListObject) As ListColumn:       Set LcFst = L.ListColumns(1):                   End Function
Function LcLas(L As ListObject) As ListColumn:       Set LcLas = L.ListColumns(L.ListColumns.Count): End Function
Function CWbLas() As Workbook:                      Set CWbLas = WbLas(Xls):                         End Function
Function WbLas(X As Excel.Application) As Workbook:  Set WbLas = X.Workbooks(X.Workbooks.Count):     End Function

Function PtRg(R As Range, Optional Wsn$, Optional Ptn$) As PivotTable
Dim Wb As Workbook: Set Wb = WbRg(R)
Dim Ws As Worksheet: Set Ws = WsAdd(Wb)
Dim A1 As Range: Set A1 = A1Ws(Ws)
Dim Pc As PivotCache: Set Pc = WbRg(R).PivotCaches.Create(xlDatabase, R.Address, Version:=6)
Dim Pt As PivotTable: Set Pt = Pc.CreatePivotTable(A1, DefaultVersion:=6)
End Function
Function PivFld(T As PivotTable, Fldn) As PivotField:       Set PivFld = T.PivotFields(Fldn):  End Function
Function PivFldCol(T As PivotTable, Coln) As PivotField: Set PivFldCol = T.RowFields(Coln):    End Function
Function PivFldRow(T As PivotTable, Rown) As PivotField: Set PivFldRow = T.ColumnFields(Rown): End Function
Function PivFldPag(T As PivotTable, Pagn) As PivotField: Set PivFldPag = T.PageFields(Pagn):   End Function

Function FmtFbtQt$(A As QueryTable)
If IsNothing(A) Then Exit Function
Dim Ty As XlCmdType, Tbl$, Cns$
With A
    Ty = .CommandType
    If Ty <> xlCmdTable Then Exit Function
    Tbl = .CommandText
    Cns = .Connection
End With
FmtFbtQt = FmtQQ("[?].[?]", ScvDtaSrc(Cns), Tbl)
End Function

Function CnoBefFstHid%(S As Worksheet)
Dim J%: For J% = 1 To MaxCno
    If CWs(S, J).Hidden Then CnoBefFstHid = J - 1: Exit Function
Next
Stop
End Function
Function DrAt(At As Range) As Variant():  DrAt = DrSq(SqRg(RgAtDown(At))):  End Function
Function DcAt(At As Range) As Variant():  DcAt = DcSq(SqRg(RgAtRight(At))): End Function
Function ValA1(R As Range):              ValA1 = RgRC(R, 1, 1).Value:       End Function

Function HasFx(Fx) As Boolean
Dim Wb As Workbook: For Each Wb In Xls.Workbooks
    If Wb.FullName = Fx Then HasFx = True: Exit Function
Next
End Function

Private Sub B_RgMoreBelow()
Dim R As Range, Act As Range, Ws As Worksheet
Set Ws = WsNw
Set R = Ws.Range("A3:B5")
Set Act = RgMoreTop(R, 1)
Debug.Print Act.Address
Stop
Debug.Print RgRR(R, 1, 2).Address
Stop
End Sub


Function RgAtDown(At As Range) As Range
Dim A1 As Range: Set A1 = A1Rg(At)
If IsEmpty(A1.Value) Then Set RgAtDown = A1: Exit Function
Dim C2&: C2 = A1.End(xlRight).Column - A1.Column + 1
Set RgAtDown = RgCRR(A1, 1, 1, C2)
End Function

Function RgAtRight(At As Range) As Range
Dim A1 As Range: Set A1 = A1Rg(At)
If IsEmpty(A1.Value) Then Set RgAtRight = A1: Exit Function
Dim R2&: R2 = A1.End(xlDown).Row - A1.Row + 1
Set RgAtRight = RgCRR(A1, 1, 1, R2)
End Function

Function DcRg(R As Range, C) As Variant(): DcRg = DcSq(SqRg(R)): End Function

Sub SwapCellVal(C1 As Range, C2 As Range)
Dim A: A = RgRC(C1, 1, 1).Value
RgRC(C1, 1, 1).Value = RgRC(C2, 1, 1).Value
RgRC(C2, 1, 1).Value = A
End Sub

Function RgSq(Sq(), At As Range) As Range
If Si(Sq) = 0 Then Set RgSq = A1Rg(At): Exit Function
Set RgSq = RgRCRC(At, 1, 1, NDrSq(Sq), NDcSq(Sq))
RgSq.Value = Sq
End Function

Function RCnoSq(A As Range) As Variant()
'Ret : #Sq-Fm:Rg-How:No ! a sq fm:@A how:no means the @ret:sq is using Rno & Cno as index.
Dim O(): O = SqRg(A)
Dim R1&, R2&, C1%, C2%
ReDim Preserve O(R1 To R2, C1 To C2)
RCnoSq = O
End Function

Function WbRg(A As Range) As Workbook
Set WbRg = WbWs(WsRg(A))
End Function

Function WsRg(A As Range) As Worksheet
Set WsRg = A.Parent
End Function

Function DrsLon(Ws As Worksheet, Lon$) As Drs: DrsLon = DrsLo(Ws.ListObjects(Lon)): End Function

Function LoLc(Lc As ListColumn) As ListObject: Set LoLc = Lc.Parent: End Function
Function DcLc(Lc As ListColumn) As Variant()
If LoLc(Lc).ListRows.Count = 0 Then Exit Function
DcLc = DcSq(SqRg(Lc.DataBodyRange))
End Function

Function DcLo(L As ListObject, C) As Variant():      DcLo = DcLc(Lc(L, C)):    End Function
Function DcStrLo(L As ListObject, C) As String(): DcStrLo = DcStrLc(Lc(L, C)): End Function

Function S12yLoP12(L As ListObject, P12$) As S12()
Dim C1, C2
If P12 = "" Then
    C1 = 1: C2 = 2
Else
    AsgT2 P12, C1, C2
End If
S12yLoP12 = S12yLo(L, C1, C2)
End Function
Function S12yLo(L As ListObject, Optional C1 = 1, Optional C2 = 2) As S12()
S12yLo = S12yAy12( _
    DcStrLc(Lc(L, C1)), _
    DcStrLc(Lc(L, C2)))
End Function
Function DcStrLc(Lc As ListColumn) As String(): DcStrLc = DcStrSq(SqRg(Lc.DataBodyRange)): End Function


Function DbTmpFx(Fx) As Database: Set DbTmpFx = DbTmpFxWny(Fx, WnyFx(Fx)): End Function

Function DbTmpFxWny(Fx, Wny$()) As Database
Dim O As Database
   Set O = DbTmp
Dim W: For Each W In Itr(Wny)
    LnkFxw O, ">" & W, Fx, CStr(W)
Next
Set DbTmpFxWny = O
End Function

Function HasWbX(X As Excel.Application, Wbn) As Boolean
Dim B As Workbook: For Each B In X.Workbooks
    If B.Name = Wbn Then HasWbX = True: Exit Function
Next
End Function

Function NoWbnX(X As Excel.Application, Wbn) As Boolean
NoWbnX = Not HasWbX(X, Wbn)
End Function

Sub MaxvFx(Fx)
Const CSub$ = CMod & "MaxvFx"
ChkFfnExi Fx, CSub, "Excel file"
MaxvWb WbFx(Fx)
End Sub
Function WbFxX(Fx, X As Excel.Application) As Workbook
Set WbFxX = X.Workbooks.Open(Fx, UpdateLinks:=False)
Minv X.Application
End Function
Function WbFx(Fx) As Workbook:          Set WbFx = WbFxX(Fx, Xls):      End Function
Function WsFxw(Fx, Wsn$) As Worksheet: Set WsFxw = WsWb(WbFx(Fx), Wsn): End Function

Private Sub B_WsnFst()
Dim Fx$
Fx = MHO.MHOMB52.FxiSalTxt
Ept = "8601"
GoSub T1
Exit Sub
T1:
    Act = WsnFst(Fx)
    C
    Return
End Sub

Private Sub B_DbTmpFx()
Dim Db As Database: Set Db = DbTmpFx(MHO.MHOMB52.FxiSalTxt)
DmpAy Tny(Db)
Db.Close
End Sub
Sub CrttFxWny(D As Database, Fx$, Wny$())
Dim W: For Each W In Itr(Wny)
    CrtTbFxW D, Fx, W
Next
End Sub
Function CrtTbFxW(D As Database, Fx$, W, Optional T$)
Dim Tbn$: Tbn = StrDft(T, W)
End Function

Function WsCdn(B As Workbook, Cdn$) As Worksheet
Dim S As Worksheet: For Each S In B.Sheets
    If S.CodeName = Cdn Then Set WsCdn = S: Exit Function
Next
End Function

Function SqcSno(N%) As Variant()
Dim O()
ReDim O(1 To 1, 1 To N)
SqcSno = O
End Function

Function SqrSno(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
SqrSno = O
End Function
Function WsS12y(A() As S12, Optional P12$ = "S1 S2") As Worksheet
Set WsS12y = WsSq(SqS12y(A, P12))
End Function

Private Sub B_S12yWsAy12()
GoTo Z
Dim A, B
Z:
    A = SySs("A B C D E")
    B = SySs("1 2 3 4 5")
    BrwS12y S12yAy12(A, B)
End Sub

Private Sub B_WbFbOupTbl()
GoTo Z
Z:
    Dim B As Workbook
    'Set W = WbFbOupTbl(WFb)
    'VisWb W
    Stop
    'W.Close False
    Set B = Nothing
End Sub

Function FnyWs(A As Worksheet) As String()
FnyWs = FnyLo(LoFst(A))
End Function

Function HasWb(X As Excel.Application, Wbn$) As Boolean
HasWb = HasItn(Xls.Workbooks, Wbn)
End Function

Function RgyLoCnoy(L As ListObject, Cnoy$()) As Range()
Dim C: For Each C In Itr(Cnoy)
    PushObj RgyLoCnoy, L.ListColumns(C).DataBodyRange
Next
End Function

Function DcyLo(L As ListObject, Ciy) As Variant()
'Fm Ciy : #DcDrs-iX-aY ! a col-ix can be a number running fm 1 or a coln.
'Ret    : #DcDrs-Ay    ! ay-of-col.  A col is ay-of-val-of-a-col.  All col has same # of ele. @@
Dim C: For Each C In Itr(Ciy)
    Dim Lc As ListColumn: Set Lc = L.ListColumns(C)
    PushI DcyLo, DcLo(L, C)
Next
End Function

Sub FilterLo(L As ListObject, Coln)
'Ret : Set filter of all L of CWs @
Dim Ws  As Worksheet:   Set Ws = CWs
Dim C$:                      C = "Mthn"
Dim Lc  As ListColumn:  Set Lc = L.ListColumns(C)
Dim OFld%:                OFld = Lc.Index
Dim Itm():                 Itm = DcLc(Lc)
Dim Patn$:                Patn = "^Ay"
Dim OSel:                 OSel = AwPatn(Itm, Patn)
Dim ORg As Range:      Set ORg = L.Range
ORg.AutoFilter Field:=OFld, Criteria1:=OSel, Operator:=xlFilterValues
End Sub


Function LoRgIn(R As Range) As ListObject
Dim R1 As Range: Set R1 = RgRC(R, 2, 1)
Dim L As ListObject: For Each L In WsRg(R).ListObjects
    If HasRg(L, R1) Then Set LoRgIn = L: Exit Function
Next
End Function
