Attribute VB_Name = "MxDta_Da_Drs_Rds"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Drs_Rds."
Type RdsDrs  ' #Reduced-Drs ! if a drs col all val are sam, mov those cols to @RdsColDic (Dic-of-coln-to-val).
    Drs As Drs       '              ! the drs aft rmv the sam val col
    CnstCol As Dictionary '        ! one entry is one col.  Key is colNm and val is colNm val.
    FldSum As Dictionary ' For all Numerice column, sum them
    DupCol As Dictionary ' Key=DupCol Val=ColWillShw
End Type

Function RdsDrs(D As Drs, ColnnSum$) As RdsDrs
If Si(D.Dy) <= 1 Then
    RdsDrs = WNoRds(D)
    Exit Function
End If
Dim O As RdsDrs
Set O.CnstCol = WDiCnst(D)
Set O.DupCol = WDiDup(D, DikyStr(O.CnstCol))
Set O.FldSum = WDiSum(D, ColnnSum)
Dim Fny$(): Fny = AwDis(SyAddAp(DikyStr(O.CnstCol), DikyStr(O.DupCol)))
O.Drs = DrsDrpDcFny(D, Fny)
RdsDrs = O
End Function
Private Function WDiSum(D As Drs, ColnnSum$) As Dictionary
'@ColnnSum: :ColnTml Field needs to have summing.  It be may *ALL, which means all fields.  If the coln not found in D.Fny, throw.
Const CSub$ = CMod & "WDiSum"
Dim SumFny$()
    SumFny = WFnySum(ColnnSum, D.Fny)
    ChkIsSupAy CSub, D.Fny, SumFny
Dim O As New Dictionary
    Dim F, Tot As Dblopt: For Each F In Itr(SumFny)
    With WColTot(D, F)
        If .Som Then
            O.Add F, .D
        End If
    End With
Next
Set WDiSum = O
End Function
Private Function WFnySum(ColnnSum$, Fny$()) As String()
If ColnnSum = "*All" Then WFnySum = Fny: Exit Function
Stop 'WFnySum = Tml(ColnnSum)
End Function

Private Function WColTot(A As Drs, C) As Dblopt
If C = StrDup("#", Len(C)) Then Exit Function ' Skip index column
Dim O#
Dim Ix%: Ix = IxEle(A.Fny, C)
Dim V, Dr: For Each Dr In Itr(A.Dy)
    If UB(Dr) >= Ix Then        ' Dr may have less field the column-@C, which is convert to *Ix
        V = Dr(Ix)
        If Not IsEmpty(V) Then
            If Not IsNumeric(V) Then Exit Function
        End If
        O = O + V
    End If
Next
WColTot = SomDbl(O)
End Function
Private Function WNoRds(D As Drs) As RdsDrs
With WNoRds
       .Drs = D
Set .CnstCol = New Dictionary
Set .DupCol = New Dictionary
Set .FldSum = New Dictionary
End With
End Function
Private Function WDiCnst(A As Drs) As Dictionary
':CnstColDi: :Di ! Key is fld nm of @A with all rec has same value.  Val is that column value
Dim NCol%: NCol = NDcDy(A.Dy)
Dim Dy(), Fny$()
Fny = A.Fny
Dy = A.Dy
Dim O As New Dictionary
Dim J%: For J = 0 To NCol - 1
    If IsCnstCol(Dy, J) Then
        O.Add Fny(J), Dy(0)(J)
    End If
Next
Set WDiCnst = O
End Function
Private Sub B_WDiDup()
Dim D As Drs: D = DrsTMdn99P
Dim CnstCol As Dictionary: Set CnstCol = WDiCnst(D)
BrwDi WDiDup(D, DikyStr(CnstCol))
End Sub
Private Function WDiDup(D As Drs, ExlCnstCny$()) As Dictionary ' duplicated column dictionary
If NoRecDrs(D) Then Set WDiDup = New Dictionary: Exit Function
Dim ChkFny$(): ChkFny = SyMinus(D.Fny, ExlCnstCny)
Set WDiDup = WDiDup1(ChkFny, D)
End Function
Private Function WDiDup1(ChkFny$(), D As Drs) As Dictionary ' Chk @ChkFny if any col are dup.
Dim DoneFny$()
Set WDiDup1 = New Dictionary
Dim U%: U = UB(ChkFny)
Dim J%: For J = 0 To UB(ChkFny)
    Dim JFld$: JFld = ChkFny(J)
    If HasEle(DoneFny, JFld) Then GoTo Nxt
    Dim I%: For I = J + 1 To U
        Dim IFld$: IFld = ChkFny(I)
        If IsDrsEqCol(D, IFld, JFld) Then
            WDiDup1.Add IFld, JFld
            PushI DoneFny, IFld
        End If
    Next
Nxt:
Next
End Function
