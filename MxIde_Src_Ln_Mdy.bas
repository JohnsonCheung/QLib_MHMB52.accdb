Attribute VB_Name = "MxIde_Src_Ln_Mdy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Ln_Mdy."
Const C_Pub$ = "Public"
Const C_Prv$ = "Private"
Const C_Frd$ = "Friend"

Function LnAddPrv$(Ln, IsPrv As Boolean)
If IsPrv Then LnAddPrv = "Private " & Ln: Exit Function
LnAddPrv = Ln
End Function

Function Shtmdyy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Pub", "Prv", "Frd", "")
Shtmdyy = X
End Function

Function Mdyy() As String()
Static X$()
If Si(X) = 0 Then X = Sy(C_Pub, C_Prv, C_Frd, "")
Mdyy = X
End Function

Function Mdy$(Ln): Mdy = PfxPfxySpc(Ln, Mdyy): End Function '#Modifier# :Nm Tm1 of a line (if it is Public | Private | Friend) otherwise *Blank

Function MdySht$(ShtMdy)
Const CSub$ = CMod & "MdyShtMdy"
Dim O$
Select Case ShtMdy
Case "": O = ""
Case "Prv": O = "Private"
Case "Frd": O = "Friend"
Case Else: Thw CSub, "ShtMdy should Blnk|Prv\Frd", "@ShtMdy with error", ShtMdy
End Select
MdySht = O
End Function
Function ShtMdy$(Mdy)
Dim O$
Select Case Mdy
Case "Public", "": O = ""
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
Case Else: O = "???"
End Select
ShtMdy = O
End Function

Function IsMdy(S) As Boolean: IsMdy = HasEle(Mdyy, S): End Function

Function RmvMdy$(Ln):               RmvMdy = LTrim(RmvPfxySpc(Ln, Mdyy)):  End Function
Function MdyPrv$(IsPrv As Boolean): MdyPrv = StrTrue(IsPrv, "Private "): End Function
