Attribute VB_Name = "MxVb_Str_Esc"
':SlashC: :Chr ! It is 1 chr.  It will combine with sfx-\.  Eg.  SlashC = 'r', it measns it will be '\r'"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Esc."
Function Hex2Asc$(Asc%)
If Asc < 16 Then
    Hex2Asc = "0" & Hex(Asc)
Else
    Hex2Asc = Hex(Asc)
End If
End Function
Function Ix4Bin2(Bin2$) As Byte
Dim O As Byte
Select Case Bin2
Case "00": O = 0
Case "01": O = 1
Case "10": O = 2
Case "11": O = 3
Case Else
    ThwImposs "Ix4Bin2"
End Select
Ix4Bin2 = O
End Function
Function Hex2$(C$)
Const CSub$ = CMod & "Hex2"
If Len(C) <> 1 Then Thw CSub, "C should have len=1", "C Len", C, Len(C)
Hex2 = Hex2Asc(Asc(C))
End Function
Function UnescChr$(S, C$): Stop: End Function
Function EscChr$(S, C$):   EscChr = EscAsc(S, Asc(C)):                    End Function
Function EscAsc$(S, A%):   EscAsc = Replace(S, Chr(A), "%" & Hex2Asc(A)): End Function 'Escaping the AscChr-A% in S$ as %HH
Function EscSqBkt$(S):   EscSqBkt = EscChrLis(S, "[]"):                   End Function
Function EscChrLis$(S, ChrLis$)
Dim O$, J: O = S
For J = 1 To Len(ChrLis)
    O = EscChr(O, Mid(ChrLis, J, 1))
Next
EscChrLis = O
End Function

Function SlashCr$(S):                        SlashCr = Slash$(S, vbCr, "r"):                End Function 'Escapeing vbCr in S.
Function EscOpnSqBkt$(S):                EscOpnSqBkt = EscChr(S, "["):                      End Function
Function UnslashCrLfTab$(S):          UnslashCrLfTab = UnslashCr(UnslashLf(UnslashTab(S))): End Function
Function EscClsSqBkt$(S):                EscClsSqBkt = EscChr(S, "]"):                      End Function
Function UnslashCrLf$(S):                UnslashCrLf = UnslashCr(UnslashLf(S)):             End Function
Function SlashAsc$(S, Asc%, C$):            SlashAsc = Slash(S, Chr(Asc), C):               End Function
Function SlashCrLf$(S):                    SlashCrLf = SlashLf(SlashCr(S)):                 End Function
Function SlashCrLfTab$(S):              SlashCrLfTab = SlashTab(SlashLf(SlashCr(S))):       End Function
Function SlashLf$(S):                        SlashLf = SlashAsc(S, 10, "n"):                End Function
Function UnslashChr$(S, C$, SlashC$):     UnslashChr = Replace(S, "\" & SlashC, C):         End Function

Function Slash$(S, C$, SlashC$) 'Escaping C$ in S by \SlashC$.  Eg C$ is vbCr and SlashC is r.
If InStr(S, "\" & SlashC) > 0 Then
    Debug.Print FmtQQ("SlashChr: Given S has \?, when Unslash, it will not match", SlashC)
    Debug.Print vbTab; QuoSq(S)
End If
Slash = Replace(S, C, "\" & SlashC)
End Function

Function UnescBackSlash$(S): UnescBackSlash = UnescChr(S, "\"):        End Function
Function EscBackSlash$(S):     EscBackSlash = EscChr(S, "\"):          End Function
Function EscCr$(S):                   EscCr = Esc(S, vbCr):            End Function
Function EscCrLf$(S):               EscCrLf = EscCr(EscLf(S)):         End Function
Function EscLf$(S):                   EscLf = EscChr(S, vbLf):         End Function
Function SlashTab$(S):             SlashTab = SlashChr(S, vbTab, "t"): End Function
Function SlashChr$(S, C$, SlashC$)
SlashChr = Replace(S, C, "\" & SlashC)
End Function
Function Esc$(S, C$):                       Esc = EscAsc(S, Asc(C)):       End Function
Function UnslashCr$(S):               UnslashCr = Replace(S, "\r", vbCr):  End Function
Function UnslashTab(S):              UnslashTab = Replace(S, "\t", vbTab): End Function
Function UnslashBackSlash$(S): UnslashBackSlash = Replace(S, "\\", "\"):   End Function
Function UnescCr$(S):                   UnescCr = Replace(S, "\r", vbCr):  End Function
Function UnescCrLf$(S):               UnescCrLf = UnescLf(UnescCr(S)):     End Function
Function UnescLf$(S): Stop: End Function
Function UnslashLf$(S):   UnslashLf = Replace(S, "\n", vbLf):                       End Function
Function UnescSpc$(S):     UnescSpc = Replace(S, "~", " "):                         End Function
Function UnescSqBkt$(S): UnescSqBkt = Replace(S, Replace(S, "\o", "["), "\c", "]"): End Function
Function UnescTab(S):      UnescTab = Replace(S, "\t", vbTab):                      End Function
