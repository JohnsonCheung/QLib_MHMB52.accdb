Attribute VB_Name = "MxDao_Rs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Rs."

Function RsSkvapC(T, ParamArray Skvap()) As Dao.Recordset
Dim Skvy(): Skvy = Skvap: Set RsSkvapC = WRsIup(CDb, T, Skvy)
End Function
Function RsSkvap(D As Database, T, ParamArray Skvap()) As Dao.Recordset
Dim Skvy(): Skvy = Skvap: Set RsSkvap = RsSkvy(D, T, Skvy)
End Function
Function RsSkvy(D As Database, T, Skvy()) As Dao.Recordset: Set RsSkvy = RsQ(D, SqlSelStarSkvy(D, T, Skvy)): End Function
Private Sub B_RsSkValApEdtC()
Dim R As Dao.Recordset: Set R = RsSkValApEdtC("Att", 1, "AAA")
Stop
End Sub
Function RsSkValApEdtC(T, ParamArray Skvap()) As Dao.Recordset
Dim Skvy(): Skvy = Skvap: Set RsSkValApEdtC = WRsIup(CDb, T, Skvy)
End Function
Function RsSkvapEdt(D As Database, T, ParamArray Skvap()) As Dao.Recordset 'return a R in edit-mode of @T if there is @Skvap else in AddNew-mode assuming @T can be insert with @Skvy only
Dim V(): V = Skvap: Set RsSkvapEdt = WRsIup(D, T, V)
End Function
Private Function WRsIup(D As Database, T, Skvy()) As Dao.Recordset 'return a R in edit-mode of @T if there is @Skvap else in AddNew-mode assuming @T can be insert with @Skvy only
Dim Q$: Q = SqlSelStarSkvy(D, T, Skvy)
If HasRecQ(D, Q) Then
    Set WRsIup = RsQ(D, Q)
    WRsIup.Edit
Else
    Set WRsIup = WRsNw(D, T, Skvy)
End If
End Function
Private Function WRsNw(D As Database, T, Skvy()) As Dao.Recordset '#Insert-R-By-Skvy# insert a new rs to @T using @Skvy, keep not Update.
Dim O As Dao.Recordset: Set O = RsTbl(D, T)
Dim J%
With O
    .AddNew
    Dim F: For Each F In FnySk(D, T)
        .Fields(F).Value = Skvy(J)  '<== Put the Skvy into the recordset
        J = J + 1
    Next
End With
Set WRsNw = O
End Function

Function AetRs(R As Dao.Recordset, Optional F = 0) As Dictionary
Set AetRs = New Dictionary
With R
    While Not .EOF
        PushAetEle AetRs, .Fields(F).Value
        .MoveNext
    Wend
End With
End Function

Sub AsgRs(R As Dao.Recordset, ParamArray OAp())
Dim F As Dao.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In R.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub

Function AvRsF(R As Dao.Recordset, Optional Fld = 0) As Variant():     AvRsF = intoDcRs(AvEmp, R, Fld): End Function
Sub BrwRs(R As Dao.Recordset):                                                 BrwDrs DrsRs(R):         End Sub
Sub BrwRec(R As Dao.Recordset):                                                BrwAy FmtRs(R):          End Sub
Function CvRs2(R) As Dao.Recordset2:                               Set CvRs2 = R:                       End Function
Function CvRs(R) As Dao.Recordset:                                  Set CvRs = R:                       End Function
Sub DltRs(R As Dao.Recordset)
With R
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub

Sub DmpRec(R As Dao.Recordset):        DmpAy FmtRs(R):       End Sub
Sub DmpRecFf(R As Dao.Recordset, FF$): DmpAy FmtRsFf(R, FF): End Sub

Function DrsRs(R As Dao.Recordset) As Drs: DrsRs = Drs(FnyRs(R), DyRs(R)): End Function

Private Sub B_DrRs()
Dim Rs As Dao.Recordset, Dy()
GoSub Z1
Exit Sub
ZZ:
    Erase Dy
    Set Rs = RsTblC("OH")
    With Rs
        While Not .EOF
            PushI Dy, DrRs(Rs)
            .MoveNext
        Wend
        .Close
    End With
    BrwDy Dy
    Return
Z1:
    Set Rs = RsTblC("Sku")
    With Rs
        Erase Dy
        While Not .EOF
            Push Dy, DrRs(Rs)
            .MoveNext
        Wend
        .Close
    End With
    BrwDy Dy
End Sub
Function DrRs(R As Dao.Recordset) As Variant():          DrRs = Itvy(R.Fields):               End Function
Function DrRsFf(R As Dao.Recordset, FF$) As Variant(): DrRsFf = DrRsFny(R.Fields, FnyFF(FF)): End Function
Function DrRsFstN(R As Dao.Recordset, N%) As Variant()
Dim Fs As Dao.Fields: Set Fs = R.Fields
Dim J%: For J = 0 To N - 1
    PushI DrRsFstN, Nz(Fs(J).Value)
Next
End Function
Function DrRsFny(R As Dao.Recordset, Fny$()) As Variant()
Dim Fs As Dao.Fields: Set Fs = R.Fields
Dim F: For Each F In Fny
    PushI DrRsFny, Nz(Fs(F).Value)
Next
End Function

Function FnyRs(R As Dao.Recordset) As String():     FnyRs = Itn(R.Fields):            End Function
Function HasRecFxQ(Fx$, Q$):                    HasRecFxQ = HasRecArs(ArsFxq(Fx, Q)): End Function

Function HasBlnkColFxw(Fx$, W$, C$) As Boolean
Dim Wh$: Wh = BeprFldIsBlnk(C)
Dim Q$: Q = SqlSelFld(C, Axtn(W), Wh)
HasBlnkColFxw = HasRecFxQ(Fx, Q)
End Function

Function HasRec(R As Dao.Recordset) As Boolean: HasRec = Not NoRec(R): End Function
Function HasRecRsFeq(R As Dao.Recordset, F, Eqval) As Boolean
If NoRec(R) Then Exit Function
With R
    .MoveFirst
    While Not .EOF
        If .Fields(F) = Eqval Then HasRecRsFeq = True: Exit Function
        .MoveNext
    Wend
End With
End Function

Sub InsRsDy(R As Dao.Recordset, Dy())
Dim Dr: For Each Dr In Itr(Dy)
    InsRs R, Dr
Next
End Sub

Sub InsRs(R As Dao.Recordset, Dr)
R.AddNew
SetRs R, Dr
R.Update
End Sub

Sub InsRsAp(R As Dao.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
InsRs R, Dr
End Sub

Sub RsDlt(R As Dao.Recordset)
With R
    If .EOF Then Exit Sub
    If .BOF Then Exit Sub
    .Delete
End With
End Sub

Function LnRs(R As Dao.Recordset, Optional Sep$ = " ")
LnRs = Join(DrRs(R), Sep)
End Function

Function JnRsFny(R As Dao.Recordset, Fny$(), Optional Sep$ = " ") As String()

End Function


Sub SetRs(R As Dao.Recordset, Dr)
Const CSub$ = CMod & "SetRs"
If Si(Dr) <> R.Fields.Count Then
    Thw CSub, "Si of R & Dr are diff", _
        "Si-R and Si-Dr R-Fny Dr", R.Fields.Count, Si(Dr), Itn(R.Fields), Dr
End If
Dim V, J%
For Each V In Dr
    If IsEmpty(V) Then
        R(J).Value = R(J).DefaultValue
    Else
        R(J).Value = V
    End If
    J = J + 1
Next
End Sub

Sub UpdRsV(R As Dao.Recordset, V)
R.Edit
R.Fields(0).Value = V
R.Update
End Sub

Sub UpdRs(R As Dao.Recordset, Dr)
R.Edit
SetRs R, Dr
R.Update
End Sub

Sub UpdRsAp(R As Dao.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
UpdRs R, Dr
End Sub

Function DcIntRs(R As Dao.Recordset, Optional Fld = 0) As Integer(): DcIntRs = intoDcRs(DcIntRs, R, Fld): End Function
Function DcLngRs(R As Dao.Recordset, Optional Fld = 0) As Long():    DcLngRs = intoDcRs(DcLngRs, R, Fld): End Function
Private Function intoDcRs(Into, R As Recordset, Optional Fld = 0):
intoDcRs = AyNw(Into)
While Not R.EOF
    PushI intoDcRs, Nz(R(Fld).Value, Empty)
    R.MoveNext
Wend
End Function

Private Sub B_SqRs()
Dim R As Dao.Recordset: Set R = RsTblC("OH")
Dim Sq(): Sq = SqRs(R, IsInlFldn:=True)
BrwSq Sq
End Sub
Function SqRs(R As Dao.Recordset, Optional IsInlFldn As Boolean) As Variant()
If True Then
    SqRs = WSqRsDim(R, IsInlFldn)
Else
    SqRs = WSqRsDy(R, IsInlFldn)
End If
End Function
Private Function WSqRsDy(R As Dao.Recordset, IsInlFldn As Boolean) As Variant()
WSqRsDy = SqDy(DyRs(R, IsInlFldn))
End Function
Private Function WSqRsDim(R As Dao.Recordset, IsInlFldn As Boolean) As Variant()
Dim NFld%: NFld = R.Fields.Count
Dim NRec&: NRec = NRecRs(R)
Dim O(): ReDim O(1 To NRec + IIf(IsInlFldn, 1, 0), 1 To NFld)
If IsInlFldn Then
    Dim IFld%: For IFld = 1 To NFld
        O(1, IFld) = R.Fields(IFld - 1).Name
    Next
End If
With R
    Dim NxtIx&: NxtIx = IIf(IsInlFldn, 2, 1)
    While Not .EOF
        W2SetSqr O, NxtIx, .Fields
        NxtIx = NxtIx + 1
        .MoveNext
    Wend
End With
WSqRsDim = O
End Function
Private Sub W2SetSqr(OSq() As Variant, R&, F As Dao.Fields)
Dim J&: J = 0
Dim I As Dao.Field: For Each I In F
    J = J + 1
    OSq(R, J) = Nz(I.Value)
Next
End Sub
