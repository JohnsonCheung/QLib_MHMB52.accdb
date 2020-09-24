Attribute VB_Name = "MxVb_Fs_Ffn_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_Op."

Function FmtFfn$(Ffn): FmtFfn = "File[" & Fn(Ffn) & "]" & vbCrLf & "Folder[" & Pth(Ffn) & "]": End Function
Function FfnTimSiStr$(Ffn$)
Dim S$: S = Format(FileLen(Ffn), "##,###,###,###")
S = AliR(S, 15)
FfnTimSiStr = "File Size/Time: [" & AliR(Format(FileLen(Ffn), "##,###,###,###"), 15) & "] [" & Format(FileDateTime(Ffn), "YYYY-MM-DD HH:MM:SS") & "]"
End Function
Function FfnLen&(Ffn)
If NoFfn(Ffn) Then
    FfnLen = -1
Else
    FfnLen = FileLen(Ffn)
End If
End Function
Function StrTimFfn$(Ffn): StrTimFfn = StrDte(FfnTim(Ffn)): End Function
Function FfnTim(Ffn) As Date
If HasFfn(Ffn) Then FfnTim = FileDateTime(Ffn)
End Function

Function IsFfnSamTimSi(Ffn1, Ffn2) As Boolean
Const CSub$ = CMod & "IsFfnSamTimSi"
ChkFfnExi Ffn1, CSub
ChkFfnExi Ffn2, CSub
If FileDateTime(Ffn1) <> FileDateTime(Ffn2) Then Exit Function
If FileLen(Ffn1) <> FileLen(Ffn2) Then Exit Function
IsFfnSamTimSi = True
End Function

Function IsFfnSamSi(Ffn1, Ffn2) As Boolean: IsFfnSamSi = FfnLen(Ffn1) = FfnLen(Ffn2): End Function

Function MsgFfnSam(A, B, Si&, Tim$, Optional Msg$) As String()
Dim O$()
Push O, "File 1   : " & A
Push O, "File 2   : " & B
Push O, "File Size: " & Si
Push O, "File Time: " & Tim
Push O, "File 1 and 2 have same size and time"
If Msg <> "" Then Push O, Msg
MsgFfnSam = O
End Function
Function IsEqFfn(A, B, Optional D As eFilCpr = eFilCprByt) As Boolean
Const CSub$ = CMod & "IsEqFfn"
ChkFfnExi A, CSub, "From File"
If A = B Then Thw CSub, "Fil A and B are eq name", "A", A
ChkFfnExi B, CSub, "To File"
If Not IsFfnSamTimSi(A, B) Then Exit Function
If D = eFilCprTimSi Then
    IsEqFfn = True
    Exit Function
End If
Dim J&, F1%, F2%
F1 = FnoRnd128(A)
F2 = FnoRnd128(B)
For J = 1 To NBlk(FfnLen(A), 128)
    If BlkFno(F1, J) <> BlkFno(F2, J) Then
        Close #F1, F2
        Exit Function
    End If
Next
Close #F1, F2
IsEqFfn = True
End Function
Function IsEqFfnStr(Ffn, S$) As Boolean
Dim L&: L = Len(S)
If FileLen(Ffn) <> L Then Exit Function
Dim J&, F%
F = FnoRnd128(Ffn)
For J = 1 To NBlk(FfnLen(Ffn), 128)
    Dim P&: P = (J - 1) * 128 + 1
    If BlkFno(F, J) <> Mid(S, P, 128) Then
        Close #F
        Exit Function
    End If
Next
Close #F
IsEqFfnStr = True
End Function
