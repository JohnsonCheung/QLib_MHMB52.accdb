Attribute VB_Name = "MxTp_IndSrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_IndSrc."
Function RmvInd(Ly$()) As String()
Dim N%: N = WNIndSpc(Ly)
Dim Fm%: Fm = N + 1
Dim L: For Each L In Itr(Ly)
    PushI RmvInd, Mid(L, Fm)
Next
End Function
Private Function WNIndSpc%(Ly$())
Dim Rx As RegExp: Set Rx = WRx
Dim L: For Each L In Itr(Ly)
    WNIndSpc = Max(WNIndSpc, PosRx(L, Rx))
Next
End Function
Private Function WRx() As RegExp
Static R As RegExp: If IsNothing(R) Then Set R = Rx("\S")
Set WRx = R
End Function

Private Sub B_SrcInd()
Dim IndtSrc$(), K$
GoSub Z
GoSub T0
Exit Sub
T0:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A 2"
    IndtSrc = XX
    Erase XX
    Ept = Sy("1", "2")
    GoTo Tst
Tst:
    Act = SrcInd(IndtSrc, K)
    C
    Return
Z:
    K = "A"
    Erase XX
    X "A Bc"
    X " 1"
    X " 2"
    X "A Bc"
    X " 1 2"
    X " 2 3"
    IndtSrc = XX
    Erase XX
    D SrcInd(IndtSrc, K)
    Return
End Sub

Function SrcInd(IndtSrc$(), Key$) As String()
Dim O$()
Dim L, Fnd As Boolean, IsNewSection As Boolean, IsfstChrSpc As Boolean, FstA%, Hit As Boolean
Const SpcAsc% = 32
For Each L In Itr(IndtSrc)
    If Left2(LTrim(L)) = "--" Then GoTo Nxt
    FstA = AscChrFst(L)
    IsNewSection = IsAscUCas(FstA)
    If IsNewSection Then
        Hit = Tm1(L) = Key
    End If
    
    IsfstChrSpc = FstA = SpcAsc
    Select Case True
    Case IsNewSection And Not Fnd And Hit: Fnd = True
    Case IsNewSection And Fnd:             SrcInd = O: Exit Function
    Case Fnd And IsfstChrSpc:              PushI O, Trim(L)
    End Select
Nxt:
Next
If Fnd Then SrcInd = O: Exit Function
End Function

