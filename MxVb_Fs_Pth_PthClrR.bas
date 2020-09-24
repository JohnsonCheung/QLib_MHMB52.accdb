Attribute VB_Name = "MxVb_Fs_Pth_PthClrR"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Pth_OpClrR."

Function IsErClrPthR(Pth) As Boolean
If NoPth(PthPar(Pth)) Then
    MsgBox "Following path not found" & vbCrLf & PthPar(Pth), vbCritical + vbDefaultButton1
    IsErClrPthR = True
    Exit Function
End If
If NoPth(Pth) Then Exit Function
DltPthR Pth
If HasPth(Pth) Then
    MsgBox "Cannot clear the following path.  May be due some file or sub-folders in the path is openned.  Close them and re-try." & vbCrLf & vbCrLf & Pth, vbDefaultButton1 + vbCritical
    IsErClrPthR = True
    Exit Function
End If
End Function

Function ResPthA$()
ResPthA = resHom & "A"
End Function

Function ResPthB$()
ResPthB = resHom & "B"
End Function

Sub CrtResPthA()
Dim P$: P = resPthzEns("A\Lvl1-A\B\C\")
WrtResl "AA", P & "AA.Txt"
WrtResl "abc", P & "ABC.Txt"
WrtResl "AA", P & "AA.txt"
WrtResl "abc", P & "ABC.Txt"
End Sub

Private Sub B_ClrPthR()
CrtResPthA
Dim T$: T = ResPthA
BrwPth T
Stop
Debug.Print IsCfmAndClrPthR(T)
End Sub

Sub ClrPth(Pth) ' Delete all file under Pth
Dim Ffn: For Each Ffn In Itr(Ffny(Pth))
    DltFfn Ffn
Next
End Sub
Sub DltPthR(Pth)
Dim F: For Each F In Itr(FfnyR(Pth))
    DltFfn F
Next
DltEmpPthR Pth
DltPthSilent Pth
End Sub

Function IsCfmAndClrPthR(Pth) As Boolean
If IsCfmClrPthR(Pth) Then
    IsCfmAndClrPthR = IsErClrPthR(Pth)
End If
End Function

Function IsCfmClrPthR(Pth) As Boolean
Const CSub$ = CMod & "IsCfmClrPthR"
If NoPth(PthPar(Pth)) Then Thw CSub, "Path not found", "Pth", Pth
If NoPth(Pth) Then IsCfmClrPthR = True: Exit Function
If MsgBox(Pth & vbCrLf & vbCrLf & "In next prompt, Input [Yes], to DELETE" & vbCrLf & "All files and folders under above path and the path itself.", vbDefaultButton1 + vbYesNo + vbQuestion) <> vbYes Then Exit Function
Dim A$: A = InputBox("Input [YES] to delete all files and folders under path in previous prompt." & vbCrLf & _
"After delete, CANNOT un-delete")
If A <> "YES" Then Exit Function
IsCfmClrPthR = True
End Function
