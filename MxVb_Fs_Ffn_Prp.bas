Attribute VB_Name = "MxVb_Fs_Ffn_Prp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_Prp."
Sub RplFfn(Ffn, ByFfn$)
BkuFfn Ffn
If DltFfnDone(Ffn) Then
    Name Ffn As ByFfn
End If
End Sub
Sub CpyPthClr(PthFm$, PthTo$)
Const CSub$ = CMod & "CpyPthClr"
ChkHasPth PthTo, CSub
DltAllPthFil PthTo
Dim Ffn$, I
For Each I In Ffny(PthFm)
    Ffn = I
    CpyFfnToPth Ffn, PthTo
Next
End Sub


Sub CpyFfnyToNxt(Ffny$())
Dim Ffn: For Each Ffn In Itr(Ffny)
    CpyFfnToNxt Ffn
Next
End Sub

Function CpyFfnToNxt$(Ffn)
Dim O$: O = FfnNxt(Ffn)
CpyFfn Ffn, O
CpyFfnToNxt = O
End Function

Sub CpyFfnyIfDif(Ffny$(), PthTo$, Optional M As eFilCpr = eFilCprByt)
Dim Ffn: For Each Ffn In Itr(Ffny)
    CpyFfnIfDif Ffn, PthTo & Fn(Ffn), M
Next
End Sub

Sub CpyFfn(FfnFm, FfnTo$, Optional OvrWrt As Boolean)
Const CSub$ = CMod & "CpyFfn"
On Error GoTo E
ChkHasPth Pth(FfnTo), CSub
ChkHasFfn FfnFm, CSub
If OvrWrt Then DltFfnIf FfnTo
Fso.CopyFile FfnFm, FfnTo
Exit Sub
E: MsgBox "Error in copying: " & Err.Description & vbCrLf & "From: " & vbCrLf & FfnFm & vbCrLf & "To:" & FfnTo
End Sub

Function CpyFfny$(Ffny$(), PthTo$, Optional OvrWrt As Boolean)
Dim P$, FfnTo$
P = PthEnsSfx(PthTo)
Dim Ffn: For Each Ffn In Ffny
    FfnTo = P & Fn(Ffn)
    CpyFfn Ffn, FfnTo, OvrWrt
Next
End Function

Sub CpyFfnIfDif(FfnFm, FfnTo$, Optional D As eFilCpr)
Const CSub$ = CMod & "CpyFfnIfDif"
If IsEqFfn(FfnFm, FfnTo, D) Then
    Dim M$: M = FmtQQ("? file", IIf(M = eFilCprByt, "EachByt", "SamTimSi"))
    Dmp MsgyFMNap(CSub, M, "FfnFm FfnTo", FfnFm, FfnTo)
    Exit Sub
End If
CpyFfn FfnFm, FfnTo, OvrWrt:=True
Dmp MsgyFMNap(CSub, "File copied", "FfnFm FfnTo", FfnFm, FfnTo)
End Sub

Sub DltFfnyIf(Ffny$())
Dim Ffn: For Each Ffn In Itr(Ffny)
    DltFfnIf Ffn
Next
End Sub
Sub DltFfn(Ffn)
Const CSub$ = CMod & "DltFfn"
On Error GoTo X
Kill Ffn
'Debug.Print "File is deleted: "; Ffn
Exit Sub
X:
Thw CSub, "Cannot delete file", "[File name] [In Folder] Er", Fn(Ffn), Pth(Ffn), Err.Description
End Sub
Sub DltFfnIf(Ffn)
If HasFfn(Ffn) Then DltFfn Ffn
End Sub
Function DltFfnIfPrompt(Ffn, Msg$) As Boolean 'Return true if error
If NoFfn(Ffn) Then Exit Function
On Error GoTo X
Kill Ffn
Exit Function
X:
MsgBox "File [" & Ffn & "] cannot be deleted, " & vbCrLf & Msg
DltFfnIfPrompt = True
End Function
Function DltFfnDone(Ffn) As Boolean
On Error GoTo X
Kill Ffn
DltFfnDone = True
Exit Function
X:
End Function

Sub MovFilUp(Pth$)
Dim Tar: For Each Tar In Itr(Fnay(Pth))
    MovFfn Tar, Pth
Next
End Sub

Sub CpyFfnUp(Ffn):                                        CpyFfnToPth Ffn, PthPar(Ffn):                 End Sub
Sub CpyFfnToPth(Ffn, PthTo$, Optional OvrWrt As Boolean): CpyFfn Ffn, FfnPthFn(PthTo, Fn(Ffn)), OvrWrt: End Sub
Sub MovFfn(Ffn, PthTo$):                                  Fso.MoveFile Ffn, PthTo:                      End Sub
