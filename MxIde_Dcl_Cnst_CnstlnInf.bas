Attribute VB_Name = "MxIde_Dcl_Cnst_CnstlnInf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Cnst_CnstlnInf."

Function CnstlnoL%(M As CodeModule, Cnstln$): CnstlnoL = LnoDclln(M, Cnstln): End Function
Function CnstlnoN%(M As CodeModule, Cnstn$)
Dim J%: For J = 1 To M.CountOfDeclarationLines
    If CnstnL(M.Lines(J, 1)) = Cnstn Then CnstlnoN = J: Exit Function
Next
End Function
Function CnstnL$(L):                              CnstnL = NmAftTm(RmvMdy(L), "Const"): End Function
Private Sub B_CnstlnyPC():                                 Brw CnstlnyPC:               End Sub
Function CnstlnyPC() As String():              CnstlnyPC = CnstlnyP(CPj):               End Function
Function CnstlnyP(P As VBProject) As String():  CnstlnyP = Cnstlny(DclP(P)):            End Function
Function Cnstlny(Src$()) As String()
Dim Ix&: For Ix = 0 To UB(Src)
    If IsLnCnst(Src(Ix)) Then PushI Cnstlny, ContlnIx(Src, Ix)
Next
End Function
Function CnstnySrc(Src$()) As String()
Dim L: For Each L In Itr(Src)
    Dim N$: N = CnstnL(L)
    Select Case N
    Case "", "CMod"
    Case Else: PushI CnstnySrc, N
    Stop
    End Select
Next
Stop
End Function
Function CnstnPubL$(L)
Dim S$: S = L
If Not IsShfPub(S) Then Exit Function
If IsShfCnst(S) Then
    CnstnPubL = TakNm(S)
End If
End Function

Private Sub B_CnstnL()
GoSub T1
Exit Sub
Dim L
T1:
    L = "Private Const AA% = 1"
    Ept = "AA"
    GoTo Tst
Tst:
    Stop
    Act = CnstnL(L)
    C
    Return
End Sub
Function CnstlnCnstn(Dcl$(), Cnstn$) As LLn
Dim J&: For J = 0 To UB(Dcl)
    Dim L$: L = Dcl(J)
    If CnstnL(L) = Cnstn Then
        L = ContlnIx(Dcl, J)
        CnstlnCnstn = LLn(J + 1, L)
        Exit Function
    End If
Next
End Function

Function CnstLLnM(M As CodeModule, Cnstn$) As LLn: CnstLLnM = CnstlnCnstn(DclM(M), Cnstn): End Function
Private Sub B_HasCnstnM()
Debug.Assert HasCnstnM(CMd, "CMod")
End Sub

Function HasCnstnM(M As CodeModule, Cnstn$) As Boolean:   HasCnstnM = CnstlnoN(M, Cnstn) = 0:      End Function
Function HasCnstnL(L, Cnstn$):                            HasCnstnL = CnstnL(L) = Cnstn:           End Function
Function HasCnstnPfx(L, CnstnPfx$) As Boolean:          HasCnstnPfx = HasPfx(CnstnL(L), CnstnPfx): End Function
Private Sub B_IsLnStrCnst()
Dim O$()
Dim L: For Each L In SrcP(CPj)
    If IsLnStrCnst(L) Then PushI O, L
Next
Brw O
End Sub
Function IsLnStrCnst(Ln) As Boolean
Dim L$: L = Ln
ShfMdy L
If Not IsShfTm(L, "Const") Then Exit Function
If ShfNm(L) = "" Then Exit Function
IsLnStrCnst = ChrFst(L) = "$"
End Function

Function IsLnCnst(L) As Boolean
Dim S$: S = L
ShfMdy S
If Not IsShfTm(S, "Const") Then Exit Function
If ShfNm(S) = "" Then Exit Function
IsLnCnst = True
End Function

Function IsLnHasCnstn(L, Cnstn$) As Boolean: IsLnHasCnstn = CnstnL(L) = Cnstn: End Function
Function IxCsntn&(Src$(), Cnstn, Optional IsPrvOnly As Boolean)
Dim O&
Dim L: For Each L In Itr(Src)
    If CnstnL(L) = Cnstn Then
        Select Case True
        Case IsPrvOnly And HasPfx(L, "Public "): IxCsntn = -1
        Case Else:                              IxCsntn = O
        End Select
        Exit Function
    End If
    O = O + 1
Next
IxCsntn = -1
End Function
