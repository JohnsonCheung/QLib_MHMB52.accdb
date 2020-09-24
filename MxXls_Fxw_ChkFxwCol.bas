Attribute VB_Name = "MxXls_Fxw_ChkFxwCol"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_ChkFxwCol."
Private Type VEuTyMisMch: F As String: ActTy As ADODB.DataTypeEnum: VTyEpt As eXlsTy: End Type ' Deriving(Ctor Ay)
Private Type VTyEpt: F As String: Ty As eXlsTy: End Type  ' Deriving(Ctor Ay)
Private Type VTyCol: F As String: Ty As ADODB.DataTypeEnum: End Type
Private Sub B_ChkFxwCol()
GoSub T1
Exit Sub
Dim Fx$, W$, CslFldn$, CslFldTy$
T1:
    Const W2a$ = "Material,Plant,Storage Location,Batch,Base Unit of Measure"
    Const W3a$ = "T       ,T    ,T               ,T    ,T                   "
    Const W2b$ = ",Unrestricted,Transit and Transfer,In Quality Insp#,Blocked,Value Unrestricted,Val# in Trans#/Tfr,Value in QualInsp#,Value BlockedStock,Value Rets Blocked"
    Const W3b$ = ",N           ,N                   ,N               ,N      ,N                 ,N                 ,N                 ,N                 ,N"
    CslFldn = W2a & W2b
    CslFldTy = W3a & W3b
    Fx = samp_mhmb52rptdta_Lo
    W = MH.MB52IO.WsnFxi
    CslFldn = MH.MB52Load.CslFldn
    CslFldTy = MH.MB52Load.CslFldTy
    GoTo Tst
Tst:
    ChkFxwCol Fx, W, CslFldn, CslFldTy
    Return
End Sub
Private Sub B_VTyColAy()
Dim A() As VTyCol: A = WTyAct(MH.MB52IO.Fxi(MH.TbOH.YmdLas), "Sheet1")
End Sub
Sub ChkFxwCol(Fx$, W$, CslFldn$, CslXlsTy$)
Dim Er$()
    Dim TyEpt() As VTyEpt: TyEpt = WTyEpt(CslFldn, CslXlsTy)
    Dim TyAct() As VTyCol: TyAct = WTyAct(Fx, W)
    Dim FnyEpt$(): FnyEpt = WFnyEpt(TyEpt)
    Dim FnyAct$(): FnyAct = WFnyAct(TyAct)
    Dim Er1$(): Er1 = WErFldMis(FnyEpt, FnyAct)
    Dim Er2$(): Er2 = WErTyMisMch_2(SyIntersect(FnyEpt, FnyAct), TyEpt, TyAct)
    Er = SyAdd(Er1, Er2): If Si(Er) = 0 Then Exit Sub
Dim O$(): O = BoxFxw(Fx, W, "Error is found in following worksheet")
ChkEry SyAdd(O, Er), "ChkFxwCol"
End Sub
Private Sub PushVEuTyMisMch(O() As VEuTyMisMch, M As VEuTyMisMch): Dim N&: N = SiVEuTyMisMch(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Sub PushVTyCol(O() As VTyCol, M As VTyCol): Dim N&: N = SiVTyCol(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Sub PushVTyEpt(O() As VTyEpt, M As VTyEpt): Dim N&: N = SiVTyEpt(O): ReDim Preserve O(N): O(N) = M: End Sub
Private Function SiVEuTyMisMch&(A() As VEuTyMisMch): On Error Resume Next: SiVEuTyMisMch = UBound(A) + 1: End Function
Private Function SiVTyCol&(A() As VTyCol): On Error Resume Next: SiVTyCol = UBound(A) + 1: End Function
Private Function SiVTyEpt&(A() As VTyEpt): On Error Resume Next: SiVTyEpt = UBound(A) + 1: End Function
Private Function UbVEuTyMisMch&(A() As VEuTyMisMch): UbVEuTyMisMch = SiVEuTyMisMch(A) - 1: End Function
Private Function UbVTyCol&(A() As VTyCol):                UbVTyCol = SiVTyCol(A) - 1:      End Function
Private Function UbVTyEpt&(A() As VTyEpt):                UbVTyEpt = SiVTyEpt(A) - 1:      End Function
Private Function VEuTyMisMch(F, ActTy As ADODB.DataTypeEnum, VTyEpt As eXlsTy) As VEuTyMisMch
With VEuTyMisMch
    .F = F
    .ActTy = ActTy
    .VTyEpt = VTyEpt
End With
End Function
Private Function VTyCol(F, Ty As ADODB.DataTypeEnum) As VTyCol
With VTyCol
    .F = F
    .Ty = Ty
End With
End Function
Private Function VTyEpt(F, Ty As eXlsTy) As VTyEpt
With VTyEpt
    .F = F
    .Ty = Ty
End With
End Function
Private Function WErFldMis(FnyEpt$(), FnyAct$()) As String()
Dim FnyMis$(): FnyMis = SyMinus(FnyEpt, FnyAct)
If Si(FnyMis) = 0 Then Exit Function
Dim O$(): O = LyUL("There are missing [" & Si(FnyMis) & "] columns in following worksheet", "-")
Dim J%: For J = 0 To UB(FnyMis)
    PushI O, vbTab & J + 1 & " [" & FnyMis(J) & "]"
Next
PushI O, Si(FnyAct) & " Worksheet columns:"
For J = 0 To UB(FnyAct)
    PushI O, vbTab & J + 1 & " [" & FnyAct(J) & "]"
Next
WErFldMis = O
End Function
Private Function WErTyMisMch_2(FnyBoth$(), TyyEpt() As VTyEpt, TyyAct() As VTyCol) As String()
WErTyMisMch_2 = W1__Ly(W2_Eu_4(FnyBoth, TyyEpt, TyyAct))
End Function
Private Function WFnyAct(A() As VTyCol) As String()
Dim J%: For J = 0 To UbVTyCol(A)
    PushI WFnyAct, A(J).F
Next
End Function
Private Function WFnyEpt(E() As VTyEpt) As String()
Dim J%: For J = 0 To UbVTyEpt(E)
    PushI WFnyEpt, E(J).F
Next
End Function
Private Function WTyAct(Fx$, W$) As VTyCol()
Dim C As ADOX.Column: For Each C In TCatTblFxw(Fx, W).T.Columns
    PushVTyCol WTyAct, VTyCol(C.Name, C.Type)
Next
End Function
Private Function WTyEpt(CslFldn$, CslFldTy$) As VTyEpt()
Dim Fny$(): Fny = AmTrim(Split(CslFldn, ","))
Dim Ty() As eXlsTy: Ty = eXlsTyyzCsl(CslFldTy)
If Si(Fny) <> Si(Ty) Then Thw "WTyEpt", "Sz(Fny)=<>Si(XlsTyAy)", "CslFldn CslFldTy", CslFldn, CslFldTy
Dim J%: For J = 0 To UB(Fny)
    PushVTyEpt WTyEpt, VTyEpt(Fny(J), Ty(J))
Next
End Function
Private Function W2_Eu_4(FnyBoth$(), TyyEpt() As VTyEpt, TyyAct() As VTyCol) As VEuTyMisMch()
Dim F: For Each F In Itr(FnyBoth)
    Dim TyEpt As eXlsTy: TyEpt = W4_TyEpt(TyyEpt, F)
    Dim TyAct As ADODB.DataTypeEnum: TyAct = W4_TyAct(TyyAct, F)
    If Not WIsEq_Ty(TyEpt, TyAct) Then
        Stop
        PushVEuTyMisMch W2_Eu_4, VEuTyMisMch(F, TyAct, TyEpt)
    End If
Next
End Function
Private Function W1__Ly(E() As VEuTyMisMch) As String()
Dim NEu%: NEu = SiVEuTyMisMch(E): If NEu = 0 Then Exit Function
W1__Ly = LyUL("There are [" & NEu & "] columns have unexpected data type", "-")
Dim J%: For J = 0 To NEu - 1
    PushI W1__Ly, W3_LnEr(J, E(J))
Next
End Function
Private Function W3_LnEr$(Ix%, E As VEuTyMisMch)
Dim F$: F = E.F
Dim NmAct$: NmAct = StrAdoTy(E.ActTy)
Dim NmEpt$: NmEpt = EnmsXlsTy(E.VTyEpt)
W3_LnEr = FmtQQ("?. Column Name[?] should be type [?] but now [?]", Ix + 1, F, NmEpt, NmAct)
End Function
Private Function WIsEq_Ty(TyEpt As eXlsTy, TyAct As ADODB.DataTypeEnum) As Boolean
Dim O As Boolean
Select Case True
Case TyEpt = eXlsTyBool: O = TyAct = adBoolean
Case TyEpt = eXlsTyDte:  O = TyAct = adDate
Case TyEpt = eXlsTyNbr:  O = TyAct = adDouble
Case TyEpt = eXlsTyTxt:  O = TyAct = adVarWChar
Case TyEpt = eXlsTyTorN: O = (TyAct = adVarWChar) Or (TyAct = adDouble)
Case Else: ThwEnm CSub, TyEpt, EnmqssXlsTy
End Select
WIsEq_Ty = O
End Function
Private Function W4_TyAct(Tyy() As VTyCol, Fldn) As ADODB.DataTypeEnum
Dim J%: For J = 0 To UbVTyCol(Tyy)
    With Tyy(J)
        If .F = Fldn Then W4_TyAct = .Ty: Exit Function
    End With
Next
ThwImposs CSub
End Function
Private Function W4_TyEpt(Tyy() As VTyEpt, Fldn) As eXlsTy
Dim J%: For J = 0 To UbVTyEpt(Tyy)
    With Tyy(J)
        If .F = Fldn Then W4_TyEpt = .Ty: Exit Function
    End With
Next
ThwImposs CSub
End Function
