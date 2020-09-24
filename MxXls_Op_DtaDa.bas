Attribute VB_Name = "MxXls_Op_DtaDa"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Op_DtaDa."
Function WbDrs(D As Drs) As Workbook:                Set WbDrs = WbRg(RgDrs(D, A1Nw)): End Function
Function RgDrs(A As Drs, At As Range) As Range:      Set RgDrs = RgSq(SqDrs(A), At):   End Function
Function LoDrs(D As Drs, At As Range) As ListObject: Set LoDrs = LoRg(RgDrs(D, At)):   End Function

Function WsDrs(A As Drs, Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = WsNw(Wsn)
Dim L As ListObject: Set L = LoDrs(A, O.Range("A1"))
Dim Lc As ListColumn: For Each Lc In L.ListColumns
    StdFmtLc Lc
Next
Set WsDrs = O
End Function
Function RgDy(Dy(), At As Range) As Range:  Set RgDy = RgSq(SqDy(Dy), At): End Function
Function WsTsy(Tsy$()) As Worksheet:       Set WsTsy = WsSq(SqTsy(Tsy)):   End Function
Function WsDy(Dy(), Optional Wsn$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = WsNw(Wsn)
RgDy Dy, A1Ws(O)
Set WsDy = O
End Function

Private Sub B_WbDs(): Maxv WbDs(sampDs).Application: End Sub
Function WsDs(D As Ds) As Workbook
Dim O As Worksheet: Set O = WbNw(D.Dsn)
A1Ws(O).Value = "*Ds " & D.Dsn
Dim At As Range: Set At = RgWsRC(O, 2, 1)
Dim BelowN&
Dim Ay() As Dt: Ay = D.Dty
Dim J&: For J = 0 To DtUB(Ay)
    Dim Dt As Dt: Dt = Ay(J)
    LoDt Dt, At
    BelowN = 2 + Si(Dt.Dy)
    Set At = CellBelow(At, BelowN)
Next
Set WsDs = O
End Function

Function LoDt(T As Dt, At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set LoDt = LoDrs(DrsDt(T), R)
RgRC(R, 0, 1).Value = T.Dtn
End Function
