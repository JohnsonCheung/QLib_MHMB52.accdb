Attribute VB_Name = "MxXls_Lo_Op_Ren"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Op_Ren."
Sub RenLoPfxFx(Fx, PfxFm$, PfxTo$)
Dim B As Workbook: Set B = WbFx(Fx)
RenLoPfxWb B, PfxFm, PfxTo
SavWb B
B.Close
End Sub
Sub RenLoPfxWb(B As Workbook, PfxFm$, PfxTo$)
Dim S As Worksheet: For Each S In B.Sheets
    RenLoPfxWs S, PfxFm, PfxTo
Next
End Sub
Sub RenLoPfxWs(S As Worksheet, PfxFm$, PfxTo$)
Dim L As ListObject: For Each L In S.ListObjects
    RenLoPfx L, PfxFm, PfxTo
Next
End Sub
Sub RenLoPfx(L As ListObject, PfxFm$, PfxTo$)
If HasPfx(L.Name, PfxFm, eCasSen) Then
    L.Name = RplPfx(L.Name, PfxFm, PfxTo)
End If
End Sub
