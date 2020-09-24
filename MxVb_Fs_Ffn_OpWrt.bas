Attribute VB_Name = "MxVb_Fs_Ffn_OpWrt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_OpWrt."
Sub WrtAy(Ay, Ft, Optional OvrWrt As Boolean)
Dim T As Scripting.TextStream: Set T = Fso.OpenTextFile(Ft, ForWriting, True)
Dim U&: U = UB(Ay)
If U = -1 Then Exit Sub
Dim J&: For J = 0 To U - 1
    T.WriteLine Ay(J)
Next
T.Write Ay(U)
T.Close
End Sub
Sub WrtStrOvr(S, Ft): DltFfnIf Ft: WrtStr S, Ft: End Sub
Sub WrtStr(S, Ft, Optional OvrWrt As Boolean)
Const CSub$ = CMod & "WrtStr"
If HasFfn(Ft) Then
    If OvrWrt Then
        DltFfn Ft
    Else
        Thw CSub, "File Exist, not over write", "Ft", Ft
    End If
End If
Dim T As Scripting.TextStream: Set T = Fso.OpenTextFile(Ft, ForWriting, True)
Dim I: For Each I In Chunky(S)
    T.Write I
Next
T.Close
End Sub

Private Sub B_Chunky()
GoSub Z
Exit Sub
Dim S, Ept$(), Act$()
Z:
    S = SrclPC
    Act = Chunky(S)
    Ass S = Jn(Act)
    Return
End Sub
Function Chunky(S, Optional Si = 10000) As String()
Dim SLen&: SLen = Len(S)
Dim UChunk&: UChunk = (SLen - 1) \ Si
Dim J&: For J = 0 To UChunk
    Dim P&, L&
    P = J * Si + 1
    If J = UChunk Then
        L = SLen - Si * (UChunk - 1)
    Else
        L = Si
    End If
    PushI Chunky, Mid(S, P, L)
Next
End Function

Function AppStr$(S, Ft)
Dim Fno%: Fno = FnoA(Ft)
Print #Fno, S;
Close #Fno
AppStr = Ft
End Function
