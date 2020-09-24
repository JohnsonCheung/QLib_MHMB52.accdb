Attribute VB_Name = "MxVb_Str_NB"
':Ff: :TmlAy #Fldn-Spc-Sep# ! a list of Fldn has no space and separated by space.
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_NB."
Function JnNBAp$(ParamArray Sap())
'Ret : :S ! ret a str by adding each ele of @StrAp one by one, if all them is <>'' else ret blank @@
Dim Av(): Av = Sap
JnNBAp = JnNB(Av)
End Function

Function JnNB$(Ay, Optional Sep$ = ""): JnNB = Join(AwNB(Ay), Sep): End Function
Sub ChkIsNB(S, Fun$)
If Not IsNB(S) Then Thw Fun, "Given-S is not NB", "Trim(S)", Trim(S)
End Sub
Function IsNB(S) As Boolean: IsNB = Trim(S) <> "": End Function

Function NBPfxDot$(IfNB$):   NBPfxDot = StrPfxIfNB(".", IfNB):   End Function
Function NBPfxSpc$(IfNB$):   NBPfxSpc = StrPfxIfNB(" ", IfNB):   End Function
Function NBSfxDot$(IfNB$):   NBSfxDot = StrPfxIfNB(IfNB, "."):   End Function
Function NBSfxSpc$(IfNB$):   NBSfxSpc = NBSfx(IfNB, " "):        End Function
Function NBPfxVBar$(IfNB$): NBPfxVBar = StrPfxIfNB(" | ", IfNB): End Function

Function StrPfxIfNB$(Pfx$, IfNB$)
If IfNB = "" Then Exit Function
StrPfxIfNB = Pfx & IfNB
End Function
Function NBSfx$(IfNB$, Sfx$)
If IfNB = "" Then Exit Function
NBSfx = IfNB & Sfx
End Function
