Attribute VB_Name = "MxVb_Dta_Sq_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Sq_Op."
Sub SetSqr(OSq, R, Dr, Optional NoTxtSngQ As Boolean)
Dim J&
If NoTxtSngQ Then
    For J = 0 To UB(Dr)
        If IsStr(Dr(J)) Then
            OSq(R, J + 1) = QuoSng(CStr(Dr(J)))
        Else
            OSq(R, J + 1) = WMz(Dr(J))
        End If
    Next
Else
    For J = 0 To UB(Dr)
        OSq(R, J + 1) = WMz(Dr(J))
    Next
End If
End Sub
Private Function WMz(V)
If IsMissing(V) Then Exit Function
WMz = V
End Function

Function SqInsRow(Sq(), Dr(), Optional Row& = 1)
Dim O(), C&, R&, NC&, NR&
NC = NDcSq(Sq)
NR = NDrSq(Sq)
ReDim O(1 To NR + 1, 1 To NC)
For R = 1 To Row - 1
    For C = 1 To NC
        O(R, C) = Sq(R, C)
    Next
Next
For C = 1 To NC
    O(Row, C) = Dr(C - 1)
Next
For R = NR To Row Step -1
    For C = 1 To NC
        O(R + 1, C) = Sq(R, C)
    Next
Next
SqInsRow = O
End Function

Sub PushSq(OSq(), Sq())
Const CSub$ = CMod & "PushSq"
Dim NR&: NR = UBound(OSq, 1) + UBound(Sq, 1)
Dim NC&: NC = UBound(OSq, 2)
Dim NC2&: NC2 = UBound(Sq, 2)
If NC <> NC2 Then Thw CSub, "NC of { OSq, Sq } are dif", "OSq-NC Sq-NC", NC, NC2
ReDim Preserve OSq(1 To NR, 1 To NC)
Dim R&, C&
For R = 1 To NC2
    For C = 1 To NC
        OSq(R + NR, C) = Sq(R, C)
    Next
Next
End Sub

Function AddSngQuoSq(Sq())
Dim NC%, C%, R&, O
O = Sq
NC = UBound(Sq, 2)
For R = 1 To UBound(Sq, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
AddSngQuoSq = O
End Function

Function IsEqSq(A, B) As Boolean
Dim NR&, NC&
NR = UBound(A, 1)
NC = UBound(A, 2)
If NR <> UBound(B, 1) Then Exit Function
If NC <> UBound(B, 2) Then Exit Function
Dim R&, C&
For R = 1 To NR
    For C = 1 To NC
        If A(R, C) <> B(R, C) Then
            Exit Function
        End If
    Next
Next
IsEqSq = True
End Function

Function TmlySq(Sq()) As String()
Dim R&: For R = 1 To NDrSq(Sq)
    Push TmlySq, TmyAy(DrSq(Sq, R))
Next
End Function
Function RgLoDta(L As ListObject) As Range
On Error Resume Next
Set RgLoDta = L.DataBodyRange
End Function
Function LoCellIn(Cell As Range) As ListObject
Dim L As ListObject: For Each L In WsRg(Cell).ListObjects
    If HasRg(L, Cell) Then
        Set LoCellIn = L
        Exit Function
    End If
Next
End Function
Function LoSq(Sq(), At As Range) As ListObject: Set LoSq = LoRg(RgSq(Sq, At)): End Function
Function LoRg(R As Range, Optional NoHdr As Boolean) As ListObject
Dim H As XlYesNoGuess: H = IIf(NoHdr, XlYesNoGuess.xlNo, xlYes)
Set LoRg = WsRg(R).ListObjects.Add(xlSrcRange, R, , H)
BdrAround R
SetLoTot LoRg
SetLoAutoFit LoRg
End Function

Function WsSq(Sq(), Optional NoHdr As Boolean) As Worksheet
Dim R As Range: Set R = RgSq(Sq, A1Nw)
Dim L As ListObject: Set L = LoRg(R, NoHdr)
Set WsSq = WsLo(L)
End Function

Function SqTranspose(Sq()) As Variant()
Dim NR&, NC&
NR = NDrSq(Sq): If NR = 0 Then Exit Function
NC = NDcSq(Sq): If NC = 0 Then Exit Function
Dim O(), J&, I&
ReDim O(1 To NC, 1 To NR)
For J = 1 To NR
    For I = 1 To NC
        O(I, J) = Sq(J, I)
    Next
Next
SqTranspose = O
End Function

Function CvDte(S, Optional Fun$)
Const CSub$ = CMod & "CvDte"
'Ret : a date fm @S if can be converted, otherwise empty and debug.print @S
On Error GoTo X
Dim O As Date: O = S
If NSsub(S, "/") <> 2 Then GoTo X ' ! one [/]-str is cv to yyyy/mm, which is not consider as a dte.
'                                       ! so use 2-[/] to treat as a dte str.
If Year(O) < 2000 Then GoTo X         ' ! year < 2000, treat it as str or not
CvDte = O
Exit Function
X: If Fun <> "" Then Inf CSub, "str[" & S & "] cannot cv to dte, emp is ret"
End Function
Private Sub B_LySqDrs()
Brw LySqDrs(DrsTMthPC)
End Sub
Function LySqDy(Dy()) As String():  LySqDy = LySq(SqDy(Dy)):          End Function
Function LySqDrs(D As Drs):        LySqDrs = LySqDy(D.Dy):            End Function
Function LySqWs(S As Worksheet):    LySqWs = LySqLo(LoFst(S)):        End Function
Function LySqLo(L As ListObject):   LySqLo = LySqRg(L.DataBodyRange): End Function

Function StrVal$(V, Optional Fun$)
':StrVal: :S #Xls-Cell-Str# ! A str coming fm xls cell
Dim T$: T = TypeName(V)
Dim O$
Select Case T
Case "Boolean", "Long", "Integer", "Date", "Currency", "Single", "Double": StrVal = V
Case "String": If IsVDblVdt(V) Then StrVal = "'" & V Else StrVal = SlashCrLfTab(V)
Case Else: If Fun <> "" Then Inf Fun, "Val-of-TypeName[" & T & "] cannot cv to :StrVal"
End Select
End Function

Function LySqRg(R As Range) As String(): LySqRg = LySq(SqRg(R)): End Function

Function IsEmpSq(Sq()) As Boolean
Dim R&: For R = 1 To UBound(Sq, 1)
    Dim C%: For C = 1 To UBound(Sq, 2)
        If Not IsEmpty(Sq(R, C)) Then Exit Function
    Next
Next
IsEmpSq = True
End Function

Function DrsSq(SqWiHdr()) As Drs
Dim Fny$(): Fny = DrStrSq(SqWiHdr, 1)
Dim Dy()
    Dim R&: For R = 2 To UBound(SqWiHdr, 1)
        PushI Dy, DrSq(SqWiHdr, R)
    Next
DrsSq = Drs(Fny, Dy)
End Function

Function DrFstSq(Sq()) As Variant()
Dim O(): ReDim O(UBound(Sq, 2) - 1)
Dim J%: For J = 0 To UBound(O)
    O(J) = Sq(1, J + 1)
Next
DrFstSq = O
End Function

Function sampSq1() As Variant()
Dim O(), R&, C&
Const NR& = 1000
Const NC& = 100
ReDim O(1 To NR, 1 To NC)
For R = 1 To NR
For C = 1 To NC
    O(R, C) = R + C
Next
Next
sampSq1 = O
End Function

Function DcFstSq(Sq()) As Variant()
Dim O(): ReDim O(UBound(Sq, 2) - 1)
Dim J&: For J = 0 To UBound(O)
    O(J) = Sq(J + 1, 1)
Next
DcFstSq = O
End Function
Function SqSampHdr() As Variant()
SqSampHdr = SqInsRow(sampSq, sampDrAToJ)
End Function
