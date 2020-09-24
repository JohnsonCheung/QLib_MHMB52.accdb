Attribute VB_Name = "MxIde_Mthln"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln."

Private Sub B_VtyFunLn()
GoSub ZZ
GoSub T1
GoSub T2
GoSub T3
GoSub T4
Exit Sub
Dim Funln, Act As TVt, Ept As TVt
T1:
    Funln = "Function MthPm(MthPm$) As MthPm"
    GoTo Tst
T2:
    Funln = "Function MthPm(MthPm$) As MthPm()"
    With Ept
        .Tyn = "MthPm"
        .IsAy = True
        .Tyc = ""
    End With
    GoTo Tst
T3:
    Funln = "Function MthPm$(MthPm$)"
    With Ept
        .Tyn = ""
        .IsAy = False
        .Tyc = "$"
    End With
    GoTo Tst
T4:
    Funln = "Function MthPm(MthPm$)"
    With Ept
        .Tyn = "Var"
        .IsAy = False
        .Tyc = ""
    End With
    GoTo Tst
ZZ:
    Dim O$()
    Dim L: For Each L In FunlnyPC
        Dim F$: 'F = RetAsVty(VtyFunln(L))
        If F = "" Then F = "."
        PushI O, F & " " & L
    Next
    Vc AySrtQ(FmtT1ry(O))
    Return
Tst:
    Act = VtyFunln(Funln)
    With Act
        Ass .Tyn = Ept.Tyn
        Ass .IsAy = Ept.IsAy
        Ass .Tyc = Ept.Tyc
    End With
    Return
End Sub
Private Function VtyFunAs(FunAs$) As TVt

End Function
Private Function VtyFunln(ContlnNoRmkFun) As TVt
Dim L$: L = RmvMth(ContlnNoRmkFun)
Dim Tyc$: Tyc = TakTyc(L)
Dim FunAs$
If Tyc = "" Then
FunAs = AftBkt(L)
End If

VtyFunln = WVty(Tyc, FunAs)
End Function
Private Function WVty(Tyc$, FunAs$) As TVt

End Function

Private Sub B_MthlnyNm()
GoSub T1
GoSub T2
Exit Sub
Dim Mthn$, Src$()
T1:
    Src = SrcMC
    Mthn = "MthlnyNm"
    Ept = Sy("Function MthlnyNm(Src$(), Mthn) As String()")
    GoTo Tst
T2:
    Src = SrcMC
    Mthn = "AA"
    Ept = Sy("Private Property Get AA$(A, B)", "Private Property Let AA(A, B, V$): End Property")
Tst:
    Act = MthlnyNm(Src, Mthn)
    C
    Return
End Sub
Private Property Get AA$(A, _
B)
End Property
Private Property Let AA(A, B, V$): End Property

Function Mi2MdMthlnPC() As String(): Mi2MdMthlnPC = Mi2MdMthlnP(CPj): End Function
Function Mi2MdMthlnP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy Mi2MdMthlnP, AmAddPfx(MthlnySrc(SrcCmp(C)), C.Name & " ")
Next
End Function

Function MthlnySrc(Src$()) As String(): MthlnySrc = Mthlny(Src):   End Function
Function MthlnyMC() As String():         MthlnyMC = MthlnyM(CMd):  End Function
Function MthlnyPC() As String():         MthlnyPC = MthlnyP(CPj):  End Function
Function MthlnyVC() As String():         MthlnyVC = MthlnyV(CVbe): End Function


Function SublnyPubPC() As String(): SublnyPubPC = SublnyPubP(CPj): End Function
Function SublnyPubP(P As VBProject) As String()
Dim L: For Each L In MthlnyP(P)
    If W3LnPubSub(L) Then PushI SublnyPubP, L
Next
End Function

Function MthlnyPubPC() As String(): MthlnyPubPC = MthnlyPubP(CPj): End Function
Function MthnlyPubP(P As VBProject) As String()
Dim L: For Each L In MthlnyP(P)
    If WIsLnPub(L) Then PushI MthnlyPubP, L
Next
End Function
Private Function WIsLnPub(L) As Boolean
Select Case Mdy(L)
Case "", "Public": WIsLnPub = True
End Select
End Function

Private Function W3LnPubSub(L) As Boolean
If WIsLnPub(L) Then
    W3LnPubSub = HasPfx(RmvPfx(L, "Public "), "Sub ")
End If
End Function



Function MthlnyV(V As VBE) As String():        MthlnyV = MthlnySrc(SrcV(V)): End Function
Function MthlnyP(P As VBProject) As String():  MthlnyP = Mthlny(SrcP(P)):    End Function
Function MthlnyM(M As CodeModule) As String(): MthlnyM = Mthlny(SrcM(M)):    End Function
Function Mthlny(Src$()) As String()
Dim Ix: For Each Ix In ItrMthix(Src)
    PushI Mthlny, Mthln(Src, Ix)
Next
End Function

Function MthlnM$(M As CodeModule, Lno&)
If Not IsLnMth(M.Lines(Lno, 1)) Then Exit Function
MthlnM = RmvVmk(StmtFst(ContlnM(M, Lno)))
End Function

Function Mthln$(Src$(), Ix)
If Not IsLnMth(Src(Ix)) Then Exit Function
Mthln = StmtFst(ContlnIx(Src, Ix))
End Function


Function MthlnyNmPC$(Mthn, Optional ShtMthTy$)
MthlnyNmPC = MthlnyNmP(CPj, Mthn, ShtMthTy)
End Function

Function MthlnyNmP$(P As VBProject, Mthn, Optional ShtMthTy$)
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy MthlnyNmP, MthlnyNmM(C.CodeModule, Mthn, ShtMthTy$)
Next
End Function

Function MthlnyNmM$(M As CodeModule, Mthn, Optional ShtMthTy$)
Dim S$(): S = SrcM(M)
Dim Ixy&(): Ixy = MthixyMthn(S, Mthn, ShtMthTy)
Dim Ix: For Each Ix In Itr(Ixy)
    PushI MthlnyNmM, Mthln(S, Ix)
Next
End Function

Function MthlnyNm(Src$(), Mthn, Optional ShtMthTy$) As String()
Dim Ix: For Each Ix In Itr(MthixyMthn(Src, Mthn, ShtMthTy))
    PushI MthlnyNm, Mthln(Src, Ix)
Next
End Function

Function MthnlyPub(Src$()) As String()
Dim L, Ix&: For Each L In Itr(Src)
    If IsLnMthPub(L) Then PushI MthnlyPub, ContlnIx(Src, Ix)
    Ix = Ix + 1
Next
End Function
