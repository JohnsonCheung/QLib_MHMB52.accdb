Attribute VB_Name = "MxDta_TreeWs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_TreeWs."
Private HomLas$
Sub SelectionChange(Target As Range)
If Target.Row = 1 Then Exit Sub
Static WIP As Boolean
If WIP Then Debug.Print "MTreeWs.SelectionChange: WIP": Exit Sub
WIP = True
Dim Ws As Worksheet: Set Ws = WsRg(Target)
Stop
EnsA1 A1Ws(Ws)
If Not IsAction(Ws) Then Exit Sub
ShwCurCol Target
ShwNxtCol Target
WIP = False
End Sub

Private Sub Change(Target As Range)
If Not IsA1(Target) Then Exit Sub
Dim Ws As Worksheet: Set Ws = WsRg(Target)
If Not IsActionWs(Ws) Then Exit Sub
Dim mA1 As Range: Set mA1 = A1Ws(Ws)
EnsA1 mA1
If NoPth(mA1.Value) Then Exit Sub
Dim Hom$: Hom = mA1.Value
If HomLas = Hom Then Exit Sub
'ShwEntHom Hom
'ShwFstHomFdr Ws
End Sub


Private Sub ShwCurCol(Cur As Range): ShwCurEnt Cur: End Sub
Private Sub ShwCurEnt(Cur As Range)
ClrCurCol Cur
Dim Pthy$(), Fnay$()
AsgEnt Pthy, Fnay, WPthCur(Cur)
WPutCurEnt Cur, Pthy, Fnay
MgeCurSubPthCol Si(Pthy)
MgeCurFnCol Si(Pthy), Si(Fnay)
End Sub
Private Function WPthCur$(Cur As Range):                WPthCur = PthEnsSfx(A1Rg(Cur).Value):                                       End Function
Private Sub WPutCurEnt(Cur As Range, Pthy$(), Fnay$()):           EntRg(Cur, Si(Pthy) + Si(Fnay)).Value = SqCol(AyAdd(Pthy, Fnay)): End Sub

Private Sub MgeCurSubPthCol(Si&)

End Sub
Private Sub MgeCurFnCol(Si1&, Si2&)

End Sub

Private Function EntRg(Cur As Range, EntCnt%) As Range
Dim Ws As Worksheet: Set Ws = WsRg(Cur)
Set EntRg = RgWsCRR(Ws, Cur.Column, 2, EntCnt + 1)
End Function

Private Sub ClrCurCol(Cur As Range)
Dim Ws As Worksheet: Set Ws = WsRg(Cur)
RgWsCRR(Ws, Cur.Column, 2, LasCno(Ws)).Delete
End Sub

Private Function CurColCC() As Range

End Function

Private Sub ShwNxtCol(Cur As Range): ShwRow Cur: End Sub
Private Sub ShwRow(Cur As Range)
Dim Ws As Worksheet: Set Ws = WsRg(Cur)
Dim R%: R = MaxR(Ws)
Dim LasR&: LasR = LasRno(Ws)
RgWsRR(Ws, 1, R).Hidden = False
RgWsRR(Ws, R + 1, LasR).Hidden = True
End Sub
Private Function MaxR%(Ws As Worksheet)
Dim J%: For J% = 1 To MaxC(Ws)
    MaxR = Max(MaxR, RgWsRC(Ws, 2, J).End(xlDown).Row - 1)
Next
End Function
Private Function MaxC%(Ws As Worksheet): MaxC = CnoBefFstHid(Ws): End Function
Private Sub EnsA1(A1 As Range)
If IsActionA1(A1) Then Exit Sub
A1.Value = "Please enter a valid path here"
Clear WsRg(A1)
End Sub
Private Sub Clear(Ws As Worksheet)
A1Ws(Ws).Activate
DltColFm Ws, 2
DltRowFm Ws, 2
HidColFm Ws, 2
HidRowFm Ws, 2
CWs(Ws, 1).AutoFit
End Sub
Private Function IsAction(Ws As Worksheet) As Boolean
IsAction = True
If IsActionWs(Ws) Then Exit Function
If IsActionA1(A1Ws(Ws)) Then Exit Function
IsAction = False
End Function
Private Function IsActionWs(Ws As Worksheet) As Boolean: IsActionWs = Ws.Name = "TreeWs": End Function

Private Function IsActionA1(A1 As Range) As Boolean
Dim V: V = A1.Value
If Not IsStr(V) Then Exit Function
IsActionA1 = HasPth(V)
End Function
