Attribute VB_Name = "MxDao_Dta_TDtaSrc_Lo"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dta_TDtaSrc_Lo."
Private Sub B_TDtaSrcLo()
Dim B As Workbook: Set B = MH.FcSamp.WbRpt
Dim Act As TDtaSrc: Act = TDtaSrcLo(B)
QuitWb B
Stop
End Sub
Function TDtaSrcLo(B As Workbook) As TDtaSrc ' Return TDtaSrc for all Lo_* in @B.  Columns of Filler* with Fml will be skipped
With TDtaSrcLo
    .Fm = TDtaSrcFm(B.FullName, "Lo")
    Dim L: For Each L In Itr(LoyTbl(B))
        PushTF .TF, WTF(CvLo(L))
    Next
End With
End Function
Private Function WTF(L As ListObject) As TF
With WTF
    .Tbn = Mid(L.Name, 4)
    .Fny = WFny(L)
End With
End Function
Private Function WFny(L As ListObject) As String() ' return only non-filler non-fml
Dim C As ListColumn: For Each C In L.ListColumns
    If HasPfx(A1Rg(C.DataBodyRange).Formula, "=") Then GoTo Nxt
    If HasPfx(C.Name, "Filler") Then GoTo Nxt
    PushI WFny, C.Name
Nxt:
Next
End Function
