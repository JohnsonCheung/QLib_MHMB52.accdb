Attribute VB_Name = "MxXls_Suodoku"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Suodoku."
Private Type YRrcc
    R1 As Byte 'all started from 1
    R2 As Byte
    C1 As Byte
    C2 As Byte
End Type
Private Type YEle9Rslt
    Ele9() As Variant
    HasSov As Boolean
End Type
Private Type YRslt
    SqWrk() As Variant
    HasSov As Boolean
End Type
Private Sub B_Sudoku_SolveWs()
Dim S As Worksheet
GoSub T0
Exit Sub
T0:
    Set S = ActiveSheet
    GoTo Tst
Tst:
    Sudoku_SolveWs S
    Return
End Sub
Sub Sudoku_RunSamp()
Dim S As Worksheet: Set S = WsNw
Dim A1 As Range: Set A1 = S.Range("A1")
X_PutSq_and_Fmt W_2SampSq, A1
W_2FmtColr A1
Sudoku_SolveWs S
Maxv S.Application
End Sub
Sub Sudoku_SolveWs(S As Worksheet)
Dim A1 As Range: Set A1 = S.Range("A1")
Dim A11 As Range: Set A11 = S.Range("A11")
X_PutSq_and_Fmt WSolveSq_3(WSq(A1)), A11
WFmtColr_ByCpy A1, A11
End Sub
Private Sub WFmtColr_ByCpy(AtFm As Range, AtTo As Range)
X_Rg99(AtFm).Copy
X_Rg99(AtTo).PasteSpecial xlPasteFormats
End Sub
Private Function WSolveSq_3(Sq() As Byte) As Byte()
Dim SqWrk(): SqWrk = W3_xSqWrk_8(Sq)
Dim SqRslt()
Dim HasSov As Boolean, J As Byte
HasSov = True
While HasSov
    HasSov = False
    J = J + 1: If J > 1000 Then Stop
    HasSov = False
    W3_xSetRslt HasSov, SqRslt, W3_Row_5(SqWrk)
    W3_xSetRslt HasSov, SqRslt, W3_Col_4(SqWrk)
    W3_xSetRslt HasSov, SqRslt, W3_Diag_6(SqWrk)
    W3_xSetRslt HasSov, SqRslt, W3_Small_7(SqWrk)
    SqWrk = SqRslt
Wend
WSolveSq_3 = W3_xSq_FmSqWrk(SqRslt)
End Function
Private Function WSq(At As Range) As Byte()
Dim O() As Byte: O = X_ToBytsq(X_Rg99(At).Value)
Dim I%, J As Byte
For J = 1 To 9
    For I = 1 To 9
        If Not IsEmpty(O(J, I)) Then
            O(J, I) = CByte(O(J, I))
        End If
    Next
Next
WSq = O
End Function
Private Function W_2SampSq() As Byte()
W_2SampSq = X_ToBytsq(SqDyByt(Av( _
Array(5, 0, 7, 6, 9, 0, 0, 0, 2), _
Array(9, 3, 0, 0, 0, 2, 7, 4, 5), _
Array(0, 0, 0, 3, 0, 7, 1, 0, 0), _
Array(0, 4, 5, 0, 6, 0, 3, 0, 8), _
Array(2, 0, 0, 4, 0, 0, 0, 0, 0), _
Array(0, 0, 0, 0, 0, 8, 1, 0, 2), _
Array(0, 0, 9, 0, 2, 0, 0, 1, 3), _
Array(3, 0, 0, 0, 0, 6, 0, 5, 7), _
Array(7, 0, 0, 1, 3, 0, 9, 8, 4))))
End Function
Private Sub W_2FmtColr(At As Range)
Dim Sq(): Sq = X_Rg99(At).Value
Dim R As Byte: For R = 1 To 9
    Dim C As Byte: For C = 1 To 9
        If Not IsEmpty(Sq(R, C)) Then
            RgRC(At, R, C).Interior.Color = 65535 'Yellow
        End If
    Next
Next
End Sub
Private Sub X_ChkSqWrkDup(SqWrk())
Static X As Boolean: If Not X Then X = True: Debug.Print "X_ChkSqWrkDup: should be removed...."
Dim J As Byte
For J = 1 To 9: X_ChkDupEle9 X_Ele9Col(SqWrk, J): Next
For J = 1 To 9: X_ChkDupEle9 X_Ele9Row(SqWrk, J): Next
For J = 1 To 9: X_ChkDupEle9 X_Ele9Small(SqWrk, J): Next
X_ChkDupEle9 X_Ele9Diag1(SqWrk)
X_ChkDupEle9 X_Ele9Diag2(SqWrk)
End Sub
Private Function W3_Col_4(SqWrk()) As YRslt
X_ChkSqWrkDup SqWrk
Dim O As YRslt: O.SqWrk = SqWrk
Dim J As Byte: For J = 1 To 9
    If J = 7 Then Stop
    With X_xEle9Rslt_2(X_Ele9Col(O.SqWrk, J))
        If .HasSov Then
            O.HasSov = True
            W4_SetSqWrk_Col O.SqWrk, J, .Ele9
        End If
    End With
Next
W3_Col_4 = O
End Function
Private Function W3_Diag_6(SqWrk()) As YRslt
Dim J As Byte
Dim O As YRslt: O.SqWrk = SqWrk
With X_xEle9Rslt_2(X_Ele9Diag1(O.SqWrk))
    If .HasSov Then
        O.HasSov = True
        W6_SetSqWrk_Diag1 O.SqWrk, .Ele9
    End If
End With
With X_xEle9Rslt_2(X_Ele9Diag2(O.SqWrk))
    If .HasSov Then
        O.HasSov = True
        W6_SetSqWrk_Diag2 O.SqWrk, .Ele9
    End If
End With
W3_Diag_6 = O
End Function
Private Function W3_Row_5(SqWrk()) As YRslt
Dim O As YRslt
O.SqWrk = SqWrk
Dim J As Byte: For J = 1 To 9
    With X_xEle9Rslt_2(X_Ele9Row(O.SqWrk, J))
        If .HasSov Then
            O.HasSov = True
            W5_SetSqWrk_Row O.SqWrk, J, .Ele9
        End If
    End With
Next
W3_Row_5 = O
End Function
Private Function W3_Small_7(SqWrk()) As YRslt
Dim O As YRslt
O.SqWrk = SqWrk
Dim J As Byte: For J = 1 To 9
    With X_xEle9Rslt_2(X_Ele9Small(O.SqWrk, J))
        If .HasSov Then
            O.HasSov = True
            W7_SetSqWrk_Small O.SqWrk, J, .Ele9
        End If
    End With
Next
W3_Small_7 = O
End Function
Private Sub W3_xSetRslt(OHasSov As Boolean, OSqWrk(), R As YRslt)
With R
    If .HasSov Then
        OHasSov = True: OSqWrk = .SqWrk
    Else
'        Stop
    End If
End With
End Sub
Private Function W3_xSq_FmSqWrk(SqWrk()) As Byte()
Dim O(1 To 9, 1 To 9) As Byte
Dim R%: For R = 1 To 9
    Dim C%: For C = 1 To 9
        If IsByt(SqWrk(R, C)) Then O(R, C) = SqWrk(R, C)
    Next
Next
W3_xSq_FmSqWrk = O
End Function
Private Function W3_xSqWrk_8(Sq() As Byte) As Variant()
Dim O(1 To 9, 1 To 9)
Dim R%: For R = 1 To 9
    Dim C%: For C = 1 To 9
        If Sq(R, C) > 0 Then
            O(R, C) = Sq(R, C)
        Else
            O(R, C) = BytyEmp
        End If
    Next
Next
W3_xSqWrk_8 = W8_SetPossByty_9(O)
End Function
Private Function X_Ele9Col(SqWrk(), Cno As Byte) As Variant()
Dim J As Byte: For J = 1 To 9
    PushI X_Ele9Col, SqWrk(J, Cno)
Next
End Function
Private Sub W4_SetSqWrk_Col(OSqWrk(), Cno As Byte, Ele9())
Dim J As Byte: For J = 1 To 9
    X_ChkCell_3 OSqWrk(J, Cno), Ele9(J - 1)
    OSqWrk(J, Cno) = Ele9(J - 1)
Next
End Sub
Private Function X_Ele9Row(SqWrk(), Row As Byte) As Variant()
Dim J As Byte: For J = 1 To 9
    PushI X_Ele9Row, SqWrk(Row, J)
Next
End Function
Private Sub W5_SetSqWrk_Row(OSqWrk(), Row As Byte, Ele9())
Dim Cno As Byte: For Cno = 1 To 9
    X_ChkCell_3 OSqWrk(Row, Cno), Ele9(Cno - 1)
    OSqWrk(Row, Cno) = Ele9(Cno - 1)
Next
End Sub
Private Sub X_ChkCell_3(Cell, Ele)
Static X As Boolean: If Not X Then X = True: Debug.Print "X_ChkCell_3: should be removed ...."
Dim CellIsByt As Boolean, CellIsByty As Boolean
Dim EleIsByt As Boolean, EleIsByty As Boolean
    CellIsByt = IsByt(Cell)
    CellIsByty = IsByty(Cell)
    EleIsByt = IsByt(Ele)
    EleIsByty = IsByty(Ele)
If Not CellIsByt And Not CellIsByty Then Stop
If Not EleIsByt And Not EleIsByty Then Stop
X3_PrintCell Cell, Ele
Select Case True
Case CellIsByt And EleIsByt: If Cell <> Ele Then Stop
Case CellIsByt:              Stop
Case EleIsByt:               If Not HasEle(Cell, Ele) Then Stop
Case Else:                   If Not IsSupAy(Cell, Ele) Then Stop
End Select
End Sub
Private Sub X3_PrintCell(Cell, Ele)
Exit Sub
Dim X
    X = Cell: GoSub P
    X = Ele: GoSub P
    Debug.Print
    Exit Sub
P:
    If IsByt(X) Then Debug.Print X, Else Debug.Print QuoBkt(JnSpc(X)),
    Return
End Sub
Private Function X_Ele9Diag1(SqWrk()) As Variant()
Dim J As Byte: For J = 1 To 9
    PushI X_Ele9Diag1, SqWrk(J, J)
Next
End Function
Private Function X_Ele9Diag2(SqWrk()) As Variant()
Dim J As Byte: For J = 1 To 9
    PushI X_Ele9Diag2, SqWrk(10 - J, 10 - J)
Next
End Function
Private Sub W6_SetSqWrk_Diag1(OSqWrk(), Ele9())
Dim J As Byte: For J = 1 To 9
    X_ChkCell_3 OSqWrk(J, J), Ele9(J - 1)
    OSqWrk(J, J) = Ele9(J - 1)
Next
End Sub
Private Sub W6_SetSqWrk_Diag2(OSqWrk(), Ele9())
Dim J As Byte: For J = 1 To 9
    X_ChkCell_3 OSqWrk(10 - J, 10 - J), Ele9(J - 1)
    OSqWrk(10 - J, 10 - J) = Ele9(J - 1)
Next
End Sub
Private Function X_Ele9Small(Sq(), J As Byte) As Variant()
Dim R As Byte, C As Byte
With X_YRrccJ(J)
For R = .R1 To .R2
    Stop 'For C = .P1 To .P2
        PushI X_Ele9Small, Sq(R, C)
    'Next
Next
End With
End Function
Private Sub W7_SetSqWrk_Small(OSqWrk(), J As Byte, Ele9())
Dim R As Byte, C As Byte
Dim I As Byte
With X_YRrccJ(J)
For R = .R1 To .R2
    Stop 'For C = .P1 To .P2
        X_ChkCell_3 OSqWrk(R, C), Ele9(I)
        OSqWrk(R, C) = Ele9(I)
        I = I + 1
    'Next
Next
End With
End Sub
Private Function W8_SetPossByty_9(SqWrk()) As Variant()
Dim O(): O = SqWrk
Dim R As Byte: For R = 1 To 9
    Dim Possible() As Byte: Possible = W9_BytyPoss(O, R)
    Dim C As Byte: For C = 1 To 9
        Dim V: V = SqWrk(R, C)
        If IsByty(V) Then
            O(R, C) = Possible
        End If
    Next
Next
W8_SetPossByty_9 = O
End Function
Private Function W9_BytyPoss(SqWrk(), R As Byte) As Byte()
Dim Certain() As Byte
Dim C%: For C = 1 To 9
    If IsByt(SqWrk(R, C)) Then PushI Certain, SqWrk(R, C)
Next
Dim Possible() As Byte
Dim J As Byte: For J = 1 To 9
    If Not HasEle(Certain, J) Then PushI Possible, J
Next
W9_BytyPoss = Possible
End Function
Private Function X_xEle9Rslt_2(Ele9()) As YEle9Rslt
Static Cnt%: Cnt = Cnt + 1
Const CSub$ = CMod & "X_xEle9Rslt_2"
If Cnt = 16 Then Stop
X_ChkDupEle9 Ele9
Dim O As YEle9Rslt
O.Ele9 = Ele9
Dim Ix As Byte: For Ix = 0 To 8
    With X2_Ele9RsltPerEle_4(O.Ele9, Ix)
        If .HasSov Then
            O.HasSov = True
            O.Ele9 = .Ele9
        End If
    End With
Next
X_xEle9Rslt_2 = O
X_ChkDupEle9 O.Ele9
End Function
Private Function X2_Ele9RsltPerEle_4(Ele9(), Ix As Byte) As YEle9Rslt
Dim E: E = Ele9(Ix)
If IsByt(E) Then Exit Function
If Not IsByty(E) Then ThwImposs CSub, "Each ele of The @Ele9 must be Byt or Byty, but now it is[" & TypeName(E) & "]"
Dim Certain() As Byte
     Certain = X4_BytyCertain(Ele9): If Si(Certain) = 0 Then ThwImposs CSub, "*BytyCertain must have some element"

Dim O As YEle9Rslt
    O.Ele9 = Ele9
    Dim M() As Byte: M = AyMinus(E, Certain)
    Dim SiBef%: SiBef = Si(E)
    Dim SiAft%: SiAft = Si(M)
    Select Case True
    Case SiBef < SiAft: ThwImposs CSub, "After AyMinus, # of element of result, is greater Ay-A"
    Case SiBef > SiAft
        O.HasSov = True
        If SiAft = 1 Then
            O.Ele9(Ix) = M(0)
            If Ix > 0 Then
                O.Ele9 = X4_Ele9RmvCertain(O.Ele9, Ix - 1, M(0))
            End If
        Else
            O.Ele9(Ix) = M
        End If
    End Select
X2_Ele9RsltPerEle_4 = O
End Function
Private Function X4_Ele9RmvCertain(Ele9(), U As Byte, Certain As Byte) As Variant()
Dim O(): O = Ele9
Dim J As Byte: For J = 0 To U
    Dim EleOld: EleOld = O(J)
    If IsByty(EleOld) Then
        If HasEle(EleOld, Certain) Then
            Dim EleNew: EleNew = AeEle(EleOld, Certain)
            If Si(EleNew) = 1 Then
                O(J) = EleNew(0)
            Else
                O(J) = EleNew
            End If
        End If
    End If
Next
X4_Ele9RmvCertain = O
End Function
Private Sub X_ChkDupEle9(Ele9())
Static X As Boolean: If Not X Then X = True: Debug.Print "X_ChkDupEle9: ... Should be removed"
Dim StrDup() As Byte
    Stop 'Dup = AwDup(X4_BytyCertain(Ele9))
    Stop 'If Si(Dup) Then Stop
End Sub
Private Sub X_PutSq_and_Fmt(Sq() As Byte, At As Range)
X_Rg99(At).Value = X_ToSq(Sq)
BdrAround RgRCRC(At, 1, 1, 3, 3)
BdrAround RgRCRC(At, 1, 4, 3, 6)
BdrAround RgRCRC(At, 1, 7, 3, 9)
BdrAround RgRCRC(At, 4, 1, 6, 3)
BdrAround RgRCRC(At, 4, 4, 6, 6)
BdrAround RgRCRC(At, 4, 7, 6, 9)
BdrAround RgRCRC(At, 7, 1, 9, 3)
BdrAround RgRCRC(At, 7, 4, 9, 6)
BdrAround RgRCRC(At, 7, 7, 9, 9)
RgCC(At, 1, 9).EntireColumn.ColumnWidth = 2
End Sub
Private Function X_Rg99(At As Range) As Range: Set X_Rg99 = RgRCRC(At, 1, 1, 9, 9): End Function
Private Function X_ToBytsq(Sq) As Byte()
Dim NRow%: NRow = UBound(Sq, 1)
Dim NCol%: NCol = UBound(Sq, 2)
Dim O() As Byte
ReDim O(1 To NRow, 1 To NCol)
Dim J%: For J = 1 To NRow
    Dim I%: For I = 1 To NCol
        If Not IsEmpty(Sq(J, I)) Then
            O(J, I) = Sq(J, I)
        End If
    Next
Next
X_ToBytsq = O
End Function
Private Function X_ToSq(Bytsq() As Byte) As Variant()
Dim NRow%: NRow = UBound(Bytsq, 1)
Dim NCol%: NCol = UBound(Bytsq, 2)
Dim O()
ReDim O(1 To NRow, 1 To NCol)
Dim J%: For J = 1 To NRow
    Dim I%: For I = 1 To NCol
        If Bytsq(J, I) > 0 Then
            O(J, I) = Bytsq(J, I)
        End If
    Next
Next
X_ToSq = O
End Function
Private Function X_YRrccJ(J As Byte) As YRrcc
Const CSub$ = CMod & "X_YRrccJ"
Dim O As YRrcc
Select Case J
Case 1: O = YRrcc(1, 3, 1, 3)
Case 2: O = YRrcc(1, 3, 4, 6)
Case 3: O = YRrcc(1, 3, 7, 9)
Case 4: O = YRrcc(4, 6, 1, 3)
Case 5: O = YRrcc(4, 6, 4, 6)
Case 6: O = YRrcc(4, 6, 7, 9)
Case 7: O = YRrcc(7, 9, 1, 3)
Case 8: O = YRrcc(7, 9, 4, 6)
Case 9: O = YRrcc(7, 9, 7, 9)
Case Else: Thw CSub, "Invalid J, should be 1 to 9", "J", J
End Select
X_YRrccJ = O
End Function
Private Function X4_BytyCertain(Ele9()) As Byte()
Dim I: For Each I In Ele9
    If IsByt(I) Then PushI X4_BytyCertain, I
Next
End Function
Private Function YRrcc(R1 As Byte, R2 As Byte, C1 As Byte, C2 As Byte) As YRrcc
With YRrcc
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
End Function
