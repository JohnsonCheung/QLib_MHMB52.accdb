Attribute VB_Name = "MxVb_Fs_Ft"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ft."
Sub DmpFt(Ft): D LinesFt(Ft): End Sub
Sub EnsFt(Ft, Optional S$)
Select Case True
Case Not HasFfn(Ft): WrtStr S, Ft: Inf CSub, "@Ft is created with @S", "@Ft Len(@S)", Ft, Len(S)
Case IsEqFfnStr(Ft, S): Inf CSub, "@Ft has same content as @S", "@Ft Len(@S)", Ft, Len(S)
Case Else
    Inf CSub, "@Ft has been ensured with Len(@S)", "@Ft Len(@S),Ft,Len(S)"
    DltFfnIf Ft
    WrtStr S, Ft, OvrWrt:=True
End Select
End Sub

Function LinesFtIf$(Ft)
If HasFfn(Ft) Then LinesFtIf = LinesFt(Ft)
End Function
Function LinesFt$(Ft)
With Fso.GetFile(Ft)
    If .Size = 0 Then Exit Function
    LinesFt = .OpenAsTextStream.ReadAll
End With
End Function

Private Sub B_FstNChrFfn()
MsgBox FstNChrFfn(MH.MB52Las.Fxi, 3)
End Sub

Function FstNChrFfn$(Ffn, N&)
Const CSub$ = CMod & "FstNChrFfn"
Dim L&: L = FileLen(Ffn): If N > L Then Thw CSub, "@Ft does not have N-Chr", "Ffn N", Ffn, N
Dim F%: F = FnoBin(Ffn)
FstNChrFfn = String(N, " ")
Get #F, , FstNChrFfn
Close #F
End Function
Function LyFtIf(Ft) As String()
If HasFfn(Ft) Then LyFtIf = LyFt(Ft)
End Function
Function LyFty(Fty$()) As String()
Dim Ft: For Each Ft In Itr(Fty)
    PushIAy LyFty, LyFt(Ft)
Next
End Function
Function Ln1Ft$(Ft)
Dim S As TextStream: Set S = Fso.GetFile(Ft).OpenAsTextStream
If Not S.AtEndOfStream Then
    Ln1Ft = S.ReadLine
    S.Close
End If
End Function

Function LyFt(Ft) As String()
If FileLen(Ft) < 100000000 Then
    LyFt = SplitCrLf(LinesFt(Ft))
    Exit Function
End If
Dim S As TextStream: Set S = Fso.GetFile(Ft).OpenAsTextStream
While Not S.AtEndOfStream
    PushI LyFt, S.ReadLine
Wend
S.Close
End Function
Function LyFtNB(Ft) As String(): LyFtNB = AwNB(LyFt(Ft)): End Function
Sub CrtFfn(Ffn)
'Do : Crt-Empty-Ffn
Close #FnoO(Ffn)
End Sub
