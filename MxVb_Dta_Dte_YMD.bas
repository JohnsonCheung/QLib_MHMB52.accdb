Attribute VB_Name = "MxVb_Dta_Dte_YMD"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_DteYMD."
Type Ymd: Y As Byte: M As Byte: D As Byte: End Type 'Deriving(Ay Ctor)
Type Ym: Y As Byte: M As Byte: End Type

Function YmYmd(A As Ymd) As Ym:     YmYmd = Ym(A.Y, A.M):                           End Function
Function YmYymm(StrYymYmM%) As Ym: YmYymm = Ym(StrYymYmM \ 100, StrYymYmM Mod 100): End Function
Function CurYM() As Ym:             CurYM = Ym(CYY, CMM):                           End Function
Function MmmYm$(M As Ym):           MmmYm = Format(DteFstYM(M), "MMM"):             End Function
Function Ym(Y As Byte, M As Byte) As Ym
Ym.Y = Y
Ym.M = M
End Function
Function IsEmpYM(A As Ym) As Boolean: IsEmpYM = Not IsBet(A.M, 1, 12): End Function
Function IsEmpYmd(A As Ymd) As Boolean
With A
Select Case True
Case .Y = 0, .M = 0, .D = 0: IsEmpYmd = True
End Select
End With
End Function
Function YmdYymmdd(YYMMDD&) As Ymd
With YmdYymmdd
    .Y = YYMMDD \ 10000
    .M = (YYMMDD - .Y * 10000&) \ 100
    .D = YYMMDD Mod 100
End With
End Function

Function YmdDte(D As Date) As Ymd
With YmdDte
    .Y = Year(D) - 2000
    .M = Month(D)
    .D = Day(D)
End With
End Function

Function DteYmd(A As Ymd) As Date: DteYmd = DateSerial(A.Y, A.M, A.D): End Function

Function RepYmd$(A As Ymd): RepYmd = FmtQQ("Ymd ? ? ?", A.Y, A.M, A.D): End Function
Function YmdAdd(A As Ymd, B As Ymd) As Ymd(): PushYmd YmdAdd, A: PushYmd YmdAdd, B: End Function
Sub PushYmdy(O() As Ymd, A() As Ymd): Dim J&: For J = 0 To UbYmd(A): PushYmd O, A(J): Next: End Sub
Sub PushYmd(O() As Ymd, M As Ymd): Dim N&: N = SiYmd(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiYmd&(A() As Ymd): On Error Resume Next: SiYmd = UBound(A) + 1: End Function
Function UbYmd&(A() As Ymd): UbYmd = SiYmd(A) - 1: End Function
Function Ymd(Y, M, D) As Ymd
With Ymd
    .Y = Y
    .M = M
    .D = D
End With
End Function
Private Sub B_YMy()
Dim YMyAsc() As Ym: YMyAsc = YMy(Ym(19, 12))
Dim YMyDes() As Ym: YMyDes = YMy(Ym(19, 12), eSrtDes)
End Sub
Function YMy(M As Ym, Optional Ord As eSrt, Optional NMth% = 15) As Ym()
Dim O() As Ym: ReDim O(0 To NMth - 1)
O(0) = M
Dim J%: For J = 1 To NMth - 1
    Select Case Ord
    Case eSrtAsc: O(J) = YmNxt(O(J - 1))
    Case eSrtDes: O(J) = YmPrv(O(J - 1))
    Case Else: Stop
    End Select
Next
YMy = O
End Function
