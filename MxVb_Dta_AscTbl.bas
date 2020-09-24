Attribute VB_Name = "MxVb_Dta_AscTbl"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_AscTbl."

Function WsAsc() As Worksheet: Set WsAsc = WsRg(RgSq(SqAscPrtNoNon, A1Nw)): End Function

Function SqAscPrtNoNon() As Variant()
SqAscPrtNoNon = AscSqRplNonPrt(SqAsc, 8)
End Function

Property Get SqAsc() As Variant()
Dim O(1 To 16, 1 To 16)
Dim I As Byte, J As Byte
For I = 0 To 15
For J = 0 To 15
    O(I + 1, J + 1) = Chr(I * 16 + J)
Next: Next
SqAsc = O
End Property

Function AscSqRplNonPrt(SqAsc(), RplByAsc%) As Variant()
Dim O(): O = SqAsc
Dim I%, J%
For I = 0 To 15
For J = 0 To 15
    If Not IsAscPrt(Asc(O(I + 1, J + 1))) Then
        O(I + 1, J + 1) = Chr(RplByAsc)
    End If
Next: Next
AscSqRplNonPrt = O
End Function

Sub BrwAscTbl()
Brw FmtAscTbl
End Sub

Sub DmpAscTbl()
DmpAy FmtAscTbl
End Sub
Function IsVdtAsc(SqAsc()) As Boolean
Select Case True
Case _
Not IsArray(SqAsc), _
UBound(SqAsc, 1) <> 16, _
LBound(SqAsc, 1) <> 1, _
UBound(SqAsc, 2) <> 16, _
LBound(SqAsc, 2) <> 1
Exit Function
End Select
IsVdtAsc = True
End Function
Property Get HexDigAy() As String()
Dim J%
For J = 0 To 15: PushI HexDigAy, Hex(J): Next
End Property

Function AscSqAddLbl(SqAsc()) As Variant()
Const CSub$ = CMod & "AscSqAddLbl"
If Not IsVdtAsc(SqAsc) Then Thw CSub, "Given SqAsc is invalid.  Vdt-SqAsc must 1-16 x 1-16"
Dim O(1 To 17, 1 To 17)
Dim R%, C%
For R = 2 To 17: For C = 2 To 17
    O(R, C) = SqAsc(R - 1, C - 1)
Next: Next
Dim A$(): A = HexDigAy
For R = 2 To 17
    O(R, 1) = A(R - 2)
Next
For C = 2 To 17
    O(1, C) = A(C - 2)
Next
O(1, 1) = " "
AscSqAddLbl = O
End Function

Function FmtAscTbl(Optional RplNonPrtByAsc% = 8) As String()
FmtAscTbl = FmtSq(AscSqAddLbl(SqAscPrtNoNon))
End Function

Property Get FmtAscSq() As String()
FmtAscSq = FmtSq(SqAscPrtNoNon)
End Property

Sub DmpAsc(S, Optional MaxLen& = 100)
Dim J&, C$
Debug.Print "Len=" & Len(S)
For J = 1 To Min(MaxLen, Len(S))
    C = Mid(S, J, 1)
    Debug.Print J, Asc(C), C
Next
End Sub

Sub DmpAscSq()
Dmp FmtAscSq
End Sub
