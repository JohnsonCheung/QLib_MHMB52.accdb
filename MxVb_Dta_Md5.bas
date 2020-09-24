Attribute VB_Name = "MxVb_Dta_Md5"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Md5."

Function HexByty$(A() As Byte)
Dim O$()
Dim I:  For Each I In Itr(A)
    PushI O, Right("0" & Hex(I), 2)
Next
HexByty = Join(O, "")
End Function
Function MD5Ft$(Ft$): MD5Ft = MD5(LinesFtIf(Ft)): End Function
Function MD5$(S$)
If S = "" Then Exit Function
Dim textBytes() As Byte: textBytes = S
MD5 = FmtByty(CvByty(WEnc.ComputeHash_2((textBytes))))
'https://stackoverflow.com/questions/40749766/how-to-check-when-a-vba-module-was-modified
'Public Function StringToMD5Hex(s As String) As String
'    Dim enc
'    Dim bytes() As Byte
'    Dim outstr As String
'    Dim pos As Integer
'    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
'    'Convert the string to a byte array and hash it
'    bytes = StrConv(s, vbFromUnicode)
'    bytes = enc.ComputeHash_2((bytes))
'    'Convert the byte array to a hex string
'    For pos = 0 To UBound(bytes)
'        outstr = outstr & LCase(Right("0" & Hex(bytes(pos)), 2))
'    Next
'    StringToMD5Hex = outstr
'    Set enc = Nothing
'End Function
End Function
Private Function WEnc() As Object
Static X As Object, Y As Boolean
If Not Y Then Y = True: Set X = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
Set WEnc = X
End Function

Function FmtByty$(A() As Byte, Optional NBytTogether% = 2)
Const CSub$ = CMod & "FmtByty"
Dim O$(), N%
N = Si(A)
If N Mod NBytTogether <> 0 Then Thw CSub, "Si-of-@Byty is not multiple of @NBytTogether", "Si-Byty NBytTogether", N, NBytTogether
Dim J%: For J = 1 To N \ NBytTogether
    PushI O, HexByty(WhByty(A, J, NBytTogether))
Next
FmtByty = JnSpc(O)
End Function

Function WhByty(A() As Byte, IthBlk%, SBlkSi%) As Byte()
Dim Offset%: Offset = (IthBlk - 1) * SBlkSi
Dim J%: For J = 0 To SBlkSi - 1
    PushI WhByty, A(Offset + J)
Next
End Function
