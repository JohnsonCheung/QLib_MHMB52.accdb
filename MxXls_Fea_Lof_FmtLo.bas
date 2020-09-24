Attribute VB_Name = "MxXls_Fea_Lof_FmtLo"
':Lof: :Ly #ListObject-Formatter# ! Each line is Ly with T1 LoflofT1nn"
':FldLikss: :Likss #Fld-Lik-SS# ! A :SS to expand a given Fny
':Ali:   :
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Fea_Lof_FmtLo."

Sub LoffmtLo(L As ListObject, Lof$()): LoffmtLoUd L, LofUd1(Lof): End Sub

Private Sub B_LoffmtLo()
GoSub T1
Exit Sub
Dim L As ListObject, Lofdta As Lofdta
T1:
    Set L = W3Lo
    Lofdta = W3Lofu
    GoTo Tst
Tst:
Stop '    FmtLo Lof L, LofUd
    Maxv L.Application
    Return
End Sub
Private Function W3Lo() As ListObject
Stop '
End Function
Private Function W3Lofu() As Lofdta
End Function

Sub LoffmtLoUd(L As ListObject, M As Lofdta)
Dim J%
With M
    SetLon L, .Lon
    For J = 0 To LofaliUB(.Ali): W2FmtAli L, .Ali(J): Next
    For J = 0 To LofBdrUB(.Bdr): W2FmtBdr L, .Bdr(J): Next
    For J = 0 To LofcorUB(.Cor): W2FmtCor L, .Cor(J): Next
    For J = 0 To LoffmlUB(.Fml): W2FmtFml L, .Fml(J): Next
    For J = 0 To UbLoffmt(.Fmt): W2FmtFmt L, .Fmt(J): Next
    For J = 0 To LoflblUB(.Lbl): W2FmtLbl L, .Lbl(J): Next
    For J = 0 To LoflvlUB(.Lvl): W2FmtLvl L, .Lvl(J): Next
    For J = 0 To LofsumUB(.Sum): W2FmtSum L, .Sum(J): Next
    For J = 0 To LofagrUB(.Agr): W2FmtTot L, .Agr(J): Next
    For J = 0 To LofwdtUB(.Wdt): W2FmtWdt L, .Wdt(J): Next
    VVFmtTit L, .Tit, M.Fny
End With
End Sub
Private Sub W2FmtSumFmTo(L As ListObject, SumFmToLn)
Dim FSum$, Ffm$, FTo$: Stop ': AsgT3R SumFmToLn, FSum, Ffm, FTo
RgLcEnt(L, FSum).Formula = FmtQQ("=Sum([?]:[?])", Ffm, FTo)
End Sub
Private Sub W2FmtAli(L As ListObject, A As Lofali): Dim F: For Each F In Itr(A.Fny): SetLcAli L, F, A.Ali: Next: End Sub
Private Sub W2FmtFml(L As ListObject, A As Loffml): RgLcEnt(L, A.Fldn).Formula = A.Fml:                          End Sub
Private Sub W2FmtLbl(L As ListObject, A As Loflbl): SetLcLbl L, A.Fldn, A.Lbl:                                   End Sub
Private Sub W2FmtSum(L As ListObject, A As Lofsum): SetLcSum L, A.SumFld, A.FmFld, A.ToFld:                      End Sub
Private Sub W2FmtBdr(L As ListObject, A As Lofbdr): Dim F: For Each F In Itr(A.Fny): SetLcBdr L, F, A.Bdr: Next: End Sub
Private Sub W2FmtLvl(L As ListObject, A As Loflvl): Dim F: For Each F In Itr(A.Fny): SetLcLvl L, F, A.Lvl: Next: End Sub
Private Sub W2FmtCor(L As ListObject, A As Lofcor): Dim F: For Each F In Itr(A.Fny): SetLcCor L, F, A.Cor: Next: End Sub
Private Sub W2FmtFmt(L As ListObject, A As Loffmt): Dim F: For Each F In Itr(A.Fny): SetLcFmt L, F, A.Fmt: Next: End Sub
Private Sub W2FmtTot(L As ListObject, A As Lofagr): Dim F: For Each F In Itr(A.Fny): SetLcTot L, F, A.Agr: Next: End Sub
Private Sub W2FmtWdt(L As ListObject, A As Lofwdt): Dim F: For Each F In Itr(A.Fny): Lc_Wdt(L, F) = A.Wdt: Next: End Sub
Private Sub B_W2FmtBdr()
Dim Ln$, Lo As ListObject, A As Lofbdr
Stop 'Set Lo = SampLo
GoSub T1
GoSub T2
Exit Sub
T1: Ln = "Left A B C": GoTo Tst
T2: Ln = "Left D E F": GoTo Tst
T3: Ln = "Right A B C": GoTo Tst
T4: Ln = "Center A B C": GoTo Tst
Tst:
    W2FmtBdr Lo, A      '<=='
    Stop
    Return
End Sub

Sub AddLoFml(L As ListObject, Coln$, Fml$)
Dim O As ListColumn
Set O = L.ListColumns.Add
O.Name = Coln
O.DataBodyRange.Formula = Fml
End Sub

Private Sub WFmtTotLnk(L As ListObject, C) 'Set HypLnkRgPr
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = L.ListColumns(C).DataBodyRange
Set Ws = WsRg(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, NRowRg(R) + 1, 1)
HypLnkRgPr R1, R2
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

Private Sub VVFmtTit(L As ListObject, A() As Loftit, Fny$())
Dim Sq(), R As Range
    Sq = WTitSq(A, Fny): If Si(Sq) = 0 Then Exit Sub
    Set R = WTitAt(L, UBound(Sq(), 1))
Set R = RgSq(Sq(), R)
WMgeTit R
BdrInside R
BdrAround R
End Sub
Private Sub B_WTitSq()
Dim Tit() As Loftit, Fny$()
'----
Dim A$(), Act(), Ept()
'TitLy
    Erase A
    Push A, "A A1 | A2 11 "
    Push A, "B B1 | B2 | B3"
    Push A, "C C1"
    Push A, "E E1"
    Tit = LofTitAy(A)

Fny = SySs("A B C D E")
Ept = WTitSq(Tit, Fny)
    SetSqr Ept, 1, SySs("A1 B1 C1 D E1")
    SetSqr Ept, 2, Array("A2 11", "B2")
    SetSqr Ept, 3, Array(Empty, "B3")
GoSub Tst
Exit Sub
'---
'Tit
    Erase A
    PushI A, "A AAA | skldf jf"
    PushI A, "B skldf|sdkfl|lskdf|slkdfj"
    PushI A, "C askdfj|sldkf"
    PushI A, "D fskldf"
    Tit = LofTitAy(A)
BrwSq WTitSq(Tit, Fny)

Exit Sub
Tst:
    Act = WTitSq(Tit, Fny)
    Ass IsEqSq(Act, Ept)
    Return
End Sub
Private Function WTitSq(A() As Loftit, Fny$()) As Variant()
Dim Dy()
    Dim F: For Each F In Fny
        PushI Dy, WDr(A, F)
    Next
WTitSq = SqTranspose(SqDy(Dy))
End Function
Private Function WDr(A() As Loftit, Fldn) As String()
Dim J%: For J = 0 To LoftitUB(A)
    If A(J).Fldn = Fldn Then
        WDr = A(J).Tit
        Exit Function
    End If
Next
'Here, if not found, just reutrn an Sy
End Function
Private Sub WMgeTit(TitRg As Range)
Dim J%
For J = 1 To NRowRg(TitRg)
    WMgeHoriTit RgR(TitRg, J)
Next
For J = 1 To TitRg.Columns.Count
    WMgeVertTit RgC(TitRg, J)
Next
End Sub
Private Sub WMgeHoriTit(TitRg As Range)
TitRg.Application.DisplayAlerts = False
Dim J%, C1%, C2%, V, LasV
LasV = RgRC(TitRg, 1, 1).Value
C1 = 1
For J = 2 To TitRg.Columns.Count
    V = RgRC(TitRg, 1, J).Value
    If V <> LasV Then
        C2 = J - 1
        If Not IsEmpty(LasV) Then
            RgRCC(TitRg, 1, C1, C2).MergeCells = True
        End If
        C1 = J
        LasV = V
    End If
Next
TitRg.Application.DisplayAlerts = True
End Sub
Private Sub WMgeVertTit(A As Range)
Dim J%: For J = NRowRg(A) To 2 Step -1
    WMgeCellAbove RgRC(A, J, 1)
Next
End Sub
Private Sub WMgeCellAbove(Cell As Range)
'If Not IsEmpty(A.Value) Then Exit Sub
If Cell.MergeCells Then Exit Sub
If Cell.Row = 1 Then Exit Sub
If RgRC(Cell, 0, 1).MergeCells Then Exit Sub
MgeRg RgCRR(Cell, 1, 0, 1)
End Sub
Private Function WTitAt(Lo As ListObject, NTitRow%) As Range: Set WTitAt = RgRC(Lo.DataBodyRange, 0 - NTitRow, 1): End Function
