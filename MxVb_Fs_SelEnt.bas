Attribute VB_Name = "MxVb_Fs_SelEnt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Sel."

Function PthSel$(Optional Pth$, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = Pth
    .Show
    If .SelectedItems.Count = 1 Then
        PthSel = PthEnsSfx(.SelectedItems(1))
    End If
End With
End Function

Private Sub B_PthSel()
GoTo Z
Z:
MsgBox FfnSel("C:\")
End Sub

Function FxSel$(Optional DftFx$, Optional SpecDes$ = "Select a Xlsx file")
FxSel = FfnSel(DftFx, "*.xlsx", SpecDes)
End Function

Function FfnSel$(Optional Ffn$, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With Application.FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    If HasFfn(Ffn) Then .InitialFileName = Ffn
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        FfnSel = .SelectedItems(1)
    End If
End With
End Function
