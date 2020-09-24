Attribute VB_Name = "MxVb_Str_Split"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Split."

Function SplitCma(S) As String():       SplitCma = Split(S, ","):  End Function
Function SplitCmaSpc(S) As String(): SplitCmaSpc = Split(S, ", "): End Function
Private Sub B_SplitPosy()
Dim S$, Posy%()
GoSub T1
Exit Sub
T1:
    S = "1234567890"
    Posy = Inty(3, 7, 10)
    Ept = Sy("12", "456", "89", "")
    GoTo Tst
Tst:
    Act = SplitPosy(S, Posy)
    C
    Return
End Sub
Function SplitPosy(S, Posy%()) As String() ' ret @Sy of si=Si(Posy)+1 by splitting @S
If Si(Posy) = 0 Then PushI SplitPosy, S: Exit Function
Dim PrvP%, L%
Dim P: For Each P In Posy
    L = P - PrvP - 1
    PushI SplitPosy, Mid(S, PrvP + 1, L)
    PrvP = P
Next
PushI SplitPosy, Mid(S, PrvP + 1)
End Function

Function SplitCrLf(S) As String()
If Len(S) > 100000000 Then
    Dim T$: T = FtTmp("SplitCrLfLargeStr")
    WrtStr S, T
    SplitCrLf = LyFt(T)
    Exit Function
End If
Dim O$: O = Replace(S, vbCr, "")
SplitCrLf = Split(O, vbLf)
End Function
Function SplitTab(S) As String():     SplitTab = Split(S, vbTab):                End Function
Function SplitDot(S) As String():     SplitDot = Split(S, "."):                  End Function
Function SplitColon(S) As String(): SplitColon = Split(S, ":"):                  End Function
Function SplitSemi(S) As String():   SplitSemi = Split(S, ";"):                  End Function
Function SplitSpc(S) As String():     SplitSpc = Split(S, " "):                  End Function
Function SplitSsl(S) As String():     SplitSsl = Split(RplDblSpc(Trim(S)), " "): End Function
Function SplitVBar(S) As String():   SplitVBar = CvSy(Split(S, "|")):            End Function

Function LyLsy(Lsy$()) As String()
Dim L: For Each L In Itr(Lsy)
    PushIAy LyLsy, SplitCrLf(L)
Next
End Function
