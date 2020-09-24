Attribute VB_Name = "MxVb_Fs_Ffn_Has"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_Has."

Sub AsgExiMis(OExi$(), OMis$(), _
Ffny$())
Dim Ffn
Erase OExi
Erase OMis
For Each Ffn In Itr(Ffny)
    If HasFfn(Ffn) Then
        PushI OExi, Ffn
    Else
        PushI OMis, Ffn
    End If
Next
End Sub

Sub ChkHasFfn(Ffn, Optional Fun$ = "ChkHasFfn")
If Not HasFfn(Ffn) Then
    Thw Fun, "File should exist", "Ffn", Ffn
End If
End Sub
Sub ChkNoFfn(Ffn)
Const CSub$ = CMod & "ChkNoFfn"
If Not NoFfn(Ffn) Then
    Thw CSub, "File should not exist", "Ffn", Ffn
End If
End Sub
Function HasFfn(Ffn) As Boolean:                    HasFfn = Fso.FileExists(Ffn):      End Function
Function NoFfn(Ffn) As Boolean:                      NoFfn = Not HasFfn(Ffn):          End Function
Function AetFfnExi(Ffny$()) As Dictionary:   Set AetFfnExi = AetAy(FfnyWhExist(Ffny)): End Function
Function AetFfnNExi(Ffny$()) As Dictionary: Set AetFfnNExi = AetAy(FfnyWhNExi(Ffny)):  End Function

Function FfnyWhExist(Ffny$()) As String()
Dim F: For Each F In Itr(Ffny)
    If HasFfn(F) Then PushI FfnyWhExist, F
Next
End Function
Function FfnyWhLen0(Ffny$()) As String()
Dim F: For Each F In Itr(Ffny)
    If HasFfn(F) Then
        If FileLen(F) = 0 Then
            PushI FfnyWhLen0, F
        End If
    Else
        PushI FfnyWhLen0, F
    End If
Next
End Function
Function FfnyWhNExi(Ffny$()) As String()
Dim F: For Each F In Itr(Ffny)
    If NoFfn(F) Then PushI FfnyWhNExi, F
Next
End Function

Sub ChkFfnExi(Ffn, Optional Fun$ = "ChkFfnExi", Optional Kd$ = "File")
If Not HasFfn(Ffn) Then RaiseMsgy MsgyFfnExi(Ffn, Fun, Kd)
End Sub
Sub ChkFfnNExi(Ffn, Optional Fun$ = "ChkFfnNExi", Optional Kd$ = "File")
If HasFfn(Ffn) Then RaiseMsgy MsgyFfnNExi(Ffn, Fun, Kd)
End Sub

Function MsgyFfnNExi(Ffn, Fun$, Optional Kd$ = "File") As String()
MsgyFfnNExi = MsgyFMNap(Fun, FmtQQ("[?] exist", Kd), "[File Pth] [File Name] [Current Dir]", Pth(Ffn), Fn(Ffn), CDir)
End Function
Function MsgyFfnExi(Ffn, Fun$, Optional Kd$ = "File") As String()
MsgyFfnExi = MsgyFMNap(Fun, FmtQQ("[?] not found", Kd), "[File Pth] [File Name] [Current Dir]", Pth(Ffn), Fn(Ffn), CDir)
End Function
