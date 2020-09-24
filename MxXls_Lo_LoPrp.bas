Attribute VB_Name = "MxXls_Lo_LoPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Op."

Sub BrwLo(L As ListObject): BrwDrs DrsLo(L): End Sub

Sub InsLcBef(L As ListObject, C$, BefCol$)
Dim Cno%: Cno = CnoLc(L, C)
RgLcEnt(L, C).Insert
Lc(L, C).Name = C
End Sub

Sub KeepFstLc(L As ListObject)
Dim J%: For J = L.ListColumns.Count To 2 Step -1
    L.ListColumns(J).Delete
Next
End Sub

Sub KeepFstLr(L As ListObject)
Dim J%: For J = L.ListRows.Count To 2 Step -1
    L.ListRows(J).Delete
Next
End Sub

Function LoPc(L As ListObject) As PivotCache
Dim O As PivotCache: Set O = WbLo(L).PivotCaches.Create(xlDatabase, L.Name, 6)
O.MissingItemsLimit = xlMissingItemsNone
Set LoPc = O
End Function

Function R1Lo&(L As ListObject, Optional InlHdr As Boolean)
If IsLoNoDta(L) Then
   R1Lo = L.ListColumns(1).Range.Row + 1
   Exit Function
End If
R1Lo = L.DataBodyRange.Row - IIf(InlHdr, 1, 0)
End Function

Function R2Lo&(L As ListObject, Optional InlTot As Boolean)
If IsLoNoDta(L) Then
   R2Lo = R1Lo(L)
   Exit Function
End If
R2Lo = L.DataBodyRange.Row + IIf(InlTot, 1, 0)
End Function

Function SqLo(L As ListObject) As Variant()
If IsNothing(L.DataBodyRange) Then Exit Function
SqLo = L.DataBodyRange.Value
End Function

Function WsLo(L As ListObject) As Worksheet: Set WsLo = L.Parent:              End Function
Function WbLo(L As ListObject) As Workbook:  Set WbLo = WbWs(WsLo(L)):         End Function
Function CnoWsLc%(L As ListObject, C):        CnoWsLc = Lc(L, C).Range.Column: End Function
Function LonTbn$(Tbn)
Select Case ChrFst(Tbn)
Case ">": LonTbn = "Lo_Inp" & RmvFst(Tbn)
Case "#": LonTbn = "Lo_Tmp" & RmvFst(Tbn)
Case "@": LonTbn = "Lo_Oup" & RmvFst(Tbn)
End Select
End Function
Function TbnLo$(L As ListObject): TbnLo = TbnLon(L.Name): End Function
Function TbnLon$(Lon$)
Select Case Left(Lon, 6)
Case "Lo_Inp": TbnLon = ">" & Mid(Lon, 7)
Case "Lo_Tmp": TbnLon = "$" & Mid(Lon, 7)
Case "Lo_Oup": TbnLon = "@" & Mid(Lon, 7)
End Select
End Function

Sub ResiLo(L As ListObject, NRow&)
Dim A1 As Range: Set A1 = A1Rg(L.Range)
Dim R2&: R2 = NRow + 1 + IIf(L.ShowTotals, 1, 0)
Dim C2%: C2 = L.ListColumns.Count
Dim R As Range: Set R = RgRCRC(A1, 1, 1, R2, C2)
L.Resize R
End Sub


Private Sub B_SetLoAutoFit()
Dim Ws As Worksheet: Set Ws = WsNw
Dim Sq(1 To 2, 1 To 2)
Sq(1, 1) = "L"
Sq(1, 2) = "B"
Sq(2, 1) = "123123"
Sq(2, 2) = String(1234, "L")
Ws.Range("A1:B2").Value = Sq
SetLoAutoFit LoWsDta(Ws)
ClsWsNoSav Ws
End Sub

Function AutoFilerLo(L As ListObject) As AutoFilter
Dim A: A = L.AutoFilter
If IsNothing(A) Then Stop
Set AutoFilerLo = A
End Function

Function CvAutoFilter(A) As AutoFilter:        Set CvAutoFilter = A:                      End Function
Function FxLo$(L As ListObject):                           FxLo = WbWs(WsLo(L)).FullName: End Function
Function Lc(L As ListObject, C) As ListColumn:           Set Lc = L.ListColumns(C):       End Function
Function RgDtaLc(L As ListObject, C, Optional InlTot As Boolean) As Range
Dim O As Range: Set O = Lc(L, C).DataBodyRange '<= If no data, it returns nothing
If IsNothing(O) Then
    Set O = RgRC(Lc(L, C).Range, 2, 1)
End If
If InlTot Then
    Set RgDtaLc = RgMoreBelow(O, 1)
Else
    Set RgDtaLc = O
End If
End Function

Function RgLcEnt(L As ListObject, C) As Range:        Set RgLcEnt = RgLc(L, C).EntireColumn: End Function
Function CnoLc%(L As ListObject, C):                        CnoLc = Lc(L, C).Index:          End Function
Function CnoyLoCC(L As ListObject, CC$) As Integer():    CnoyLoCC = CnoyLoFny(L, FnyFF(CC)): End Function
Function CnoyLoFny(L As ListObject, Fny$()) As Integer()
Dim C: For Each C In Itr(Fny)
    PushI CnoyLoFny, CnoLc(L, C)
Next
End Function

Function LcAft(L As ListObject, AftC) As ListColumn
Dim C%: C = CnoLc(L, AftC) + 1
Set LcAft = Lc(L, C)
End Function

Function RgLcInl(L As ListObject, C, Optional InlTot As Boolean, Optional InlHdr As Boolean) As Range
Dim R As Range
Set R = L.ListColumns(C).DataBodyRange
If Not InlTot And Not InlHdr Then
    Set RgLcInl = R
    Exit Function
End If
If InlTot Then Set RgLcInl = RgMoreBelow(R, 1)
If InlHdr Then Set RgLcInl = RgMoreTop(R, 1)
End Function
Function RgLcAft(L As ListObject, Aft) As Range: Set RgLcAft = Lc(L, CnoLc(L, Aft) + 1).Range: End Function
Function RgLcDta(L As ListObject, C) As Range:   Set RgLcDta = Lc(L, C).DataBodyRange:         End Function
Function RgLc(L As ListObject, C) As Range:         Set RgLc = Lc(L, C).Range:                 End Function

Function LoRno&(L As ListObject, Rg As Range)
':LoRno :Row-No #Listobject-Row-No# ! Fm 1-L.ListRows.Count, 0 will ix not found
If Not HasRg(L, Rg) Then Exit Function
LoRno = Rg.Row - L.DataBodyRange.Row + 1
End Function
Function HasLoRg(Rg As Range) As ListObject
Dim Rc As TRc: Rc = TRcRg(Rg)
Dim L As ListObject: For Each L In WsRg(Rg).ListObjects
    If HasTRc(RrccLo(L), Rc) Then Set HasLoRg = L: Exit Function
Next
End Function
Function HasRg(L As ListObject, Rg As Range) As Boolean:  HasRg = HasTRc(RrccLo(L), TRcRg(Rg)): End Function 'Ret :True if @L-data-range has the @Rg-A1
Function RrccLo(L As ListObject) As Rrcc:                RrccLo = RrccRg(L.DataBodyRange):      End Function
Function DyLo(L As ListObject) As Variant():               DyLo = DySq(SqLo(L)):                End Function
Function RnoLasLo(L As ListObject) As Range
Const CSub$ = CMod & "RnoLasLo"
If L.ListRows.Count = 0 Then Thw CSub, "There is no RnoLasLo in L", "Lon Wsn", L.Name, WsLo(L).Name
Set RnoLasLo = L.ListRows(L.ListRows.Count).Range
End Function

Function CellLasRow(L As ListObject, C) As Range
Dim Ix%: Ix = L.ListColumns(C).Index
Set CellLasRow = RgRC(RnoLasLo(L), 1, L.ListColumns(C).Index)
End Function

Function RgLoP12(L As ListObject, P12$) As Range
Dim A%, B%: AsgCix FnyLo(L), P12, A, B
Set RgLoP12 = RgCC(L.DataBodyRange, A, B)
End Function
Function WsnLo$(A As ListObject):                         WsnLo = WsLo(A).Name:         End Function
Function CellLcHdr(L As ListObject, C$) As Range: Set CellLcHdr = Lc(L, C).Range(1, 1): End Function

Function RgLoCC(L As ListObject, C1, C2, Optional InlTot As Boolean, Optional InlHdr As Boolean) As Range
Dim R1&, R2&, CA%, CB%
R1 = R1Lo(L, InlHdr)
R2 = R2Lo(L, InlTot)
CA = CnoWsLc(L, C1)
CB = CnoWsLc(L, C2)
Set RgLoCC = RgWsRCRC(WsLo(L), R1, CA, R2, CB)
End Function
