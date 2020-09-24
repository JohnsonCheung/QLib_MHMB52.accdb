Attribute VB_Name = "MxIde_Mth_Slm_zIntl_Slmb"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Slmb_SlmFun."
Private Sub B_SlmbyPC():                           BrwLyy SlmbyPC:             End Sub
Function SlmbyPC() As Variant():         SlmbyPC = SlmbySrc(SrcPC):            End Function
Function SlmbySrc(Src$()) As Variant(): SlmbySrc = AyyBeiy(Src, BeiySlm(Src)): End Function
Function BeiySlm(Src$()) As Bei()
Dim Booly() As Boolean
    Dim L: For Each L In Itr(Src)
        PushI Booly, IsLnSlm(L)
    Next
    BeiySlm = BeiyBooly(Booly)
End Function

Function IsLnSlm(L) As Boolean
Dim T$: T = MthTyLn(L): If T = "" Then Exit Function
Select Case T
Case "Function": If Not HasSsub(L, "End Function") Then Exit Function
                 Dim N$: N = MthnL(L)
                 Dim S$(): S = Stmty(L)
                 If Si(S) <> 3 Then Exit Function
                 If HasSsub(S(1), N & " = ") Then IsLnSlm = True '<==
Case "Sub":      If Not HasSsub(L, "End Sub", eCasSen) Then Exit Function
                 If HasSsub(L, "Sub Push") Then Exit Function
                 IsLnSlm = True  '<==
Case "Property Get", "Property Let", "Property Set":
                 IsLnSlm = HasSsub(L, "End Property") '<==
Case Else: Stop
End Select
End Function

Private Sub B_IsLnSlm()
'GoSub T1
GoSub T2
GoSub ZZ1
Exit Sub
Dim L, O$(), LySub$(), LyFun$()
ZZ:
    GoSub Fnd_O
    VcAy O
ZZ1:
    GoSub Fnd_LyFun_and_LySub
    LyFun = SlmbAli(LyFun)
    LySub = SlmbAli(LySub)
    BrwLsy Sy(JnCrLf(LyFun), JnCrLf(LySub))
    Return
T1:
    L = "Private Const ShwAllSql$ = SelSql & WhAll & OrdBy"
    Ept = False
    GoTo Tst
T2:
    L = "Function SmsEleFldSI&(A() As SmsEleFld): On Error Resume Next: SmsEleFldSi = UBound(A) + 1: End Function"
    Ept = False
    GoTo Tst
Tst:
    Act = IsLnSlm(L)
    C
    Return
Fnd_O:
    Erase O
    For Each L In SrcPC
        If IsLnSlm(L) Then PushI O, L
    Next
    Return
Fnd_LyFun_and_LySub:
    GoSub Fnd_O
    Erase LyFun, LySub
    For Each L In O
        Select Case True
        Case HasPfx(L, "Private Sub "), HasPfx(L, "Sub "): PushI LySub, L
        Case Else: PushI LyFun, L
        End Select
    Next
    Return
End Sub
