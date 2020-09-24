Attribute VB_Name = "MxIde_Mthn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn."

Function MthnL$(Ln)
Dim L$: L = RmvMdy(Ln)
If ShfMtht(L) = "" Then Exit Function
MthnL = TakNm(L)
End Function

Function PrpnL$(Ln)
Dim L$: L = RmvMdy(Ln)
If ShfMthKd(L) <> "Property" Then Exit Function
PrpnL = TakNm(L)
End Function
Function PrpnPubL$(Ln)
Dim L$: L = Ln
If Not IsShfPub(L) Then Exit Function
If ShfMthKd(L) <> "Property" Then Exit Function
PrpnPubL = TakNm(L)
End Function

Private Sub B_MthnL()
GoTo Z
Dim A$
A = "Function MthnL(A)": Ept = "Mthn.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = MthnL(A)
    C
    Return
Z:
    Dim O$(), L
    For Each L In SrcV(CVbe)
        PushNB O, MthnL(CStr(L))
    Next
    Brw O
End Sub

Function MthnLno$(M As CodeModule, Lno&)
Dim K As vbext_ProcKind
MthnLno = M.ProcOfLine(Lno, K)
End Function
Function HasMthnS(Src$(), Mthn, Optional ShtMthTy$) As Boolean
Dim L: For Each L In Itr(Src)
    With TMthL(L)
        If .Mthn = Mthn Then
            If HitOptEq(.ShtTy, ShtMthTy) Then
                HasMthnS = True
                Exit Function
            End If
            Debug.Print FmtQQ("HasMthnM: Ln has Mthn[?] but not hit given ShtMthTy[?].  Act ShtMthTy=[?]", Mthn, ShtMthTy, .ShtTy)
        End If
    End With
Next
End Function

Function HasMthnM(M As CodeModule, Nm, Optional ShtMthTy$) As Boolean: HasMthnM = HasMthnS(SrcM(M), Nm, ShtMthTy): End Function
