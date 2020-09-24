Attribute VB_Name = "MxVb_Str_Crypto"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Crypto."
Private Sub B_Asmy()
Dim O
O = Asmy
Stop
End Sub
Function Asmy() As Object()
' function Get-Assemblies { [System.AppDomain]::CurrentDomain.GetAssemblies() }
Dim AppDomain, CurDomain
Set AppDomain = CreateObject("System.AppDomain")
Set CurDomain = AppDomain.CurrentDomain
Stop
End Function
Private Sub B_GetObject()
Dim A As Excel.Application
Set A = GetObject(, "Excel.Application")
Stop
A.Workbooks.Add
A.Visible = False 'Must have workbook open to allow Visible has effect
Dim B As Excel.Application
Set B = GetObject(, "Excel.Application")
B.Workbooks.Add
B.Visible = False 'Must have workbook open to allow Visible has effect
Debug.Print ObjPtr(A), ObjPtr(B)
Stop
Stop
End Sub
Function B64Xml$(S$)
Dim B() As Byte: B = S
  'Ref: http://stackoverflow.com/questions/1118947/converting-binary-file-to-base64-string
  With CreateObject("MSXML2.DOMDocument")
    .LoadXML "<root />"
    .DocumentElement.DataType = "bin.base64"
    .DocumentElement.nodeTypedValue = B
    B64Xml = Replace(.DocumentElement.Text, vbLf, "")
  End With
End Function

Function StrB64zXml$(B64$)
Dim B() As Byte: B = B64
  'Ref: http://stackoverflow.com/questions/1118947/converting-binary-file-to-base64-string
  With CreateObject("MSXML2.DOMDocument")
    .LoadXML "<root />"
    .DocumentElement.DataType = "bin.Hex"
    .DocumentElement.nodeTypedValue = B
    StrB64zXml = Replace(.DocumentElement.Text, vbLf, "")
  End With
End Function

Private Sub B_SHA256()
'Requires a reference to mscorlib 4.0 64-bit, which is part of the .Net Framework 4.0
GoTo Tst1
Exit Sub
Tst1:
    Dim A() As Byte
    Dim Text As Object
    Dim SHA256 As Object
        A = CreateObject("System.Text.UTF8Encoding").GetBytes_4("abcd")
        Set Text = CreateObject("System.Text.UTF8Encoding")
        Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
        
        If True Then
            Dim bytes
            Dim Hash$ ' originally it is [Dim Hash]
            bytes = Text.GetBytes_4("mypassword")
            Hash = SHA256.ComputeHash_2((bytes)) ' Single brackket quote is not OK
            Debug.Print StrB64zXml(Hash)
        Else
            Debug.Print StrB64zXml(SHA256.ComputeHash_2((Text.GetBytes_4("mypassword"))))
        End If
        Stop
    ShwDbg
    Stop
    Return
End Sub

Private Sub B_SHA512()
'64-bit MS Access VBA code to calculate an SHA-512 or SHA-256 hash in VBA.  This requires a VBA reference to the .Net Framework 4.0 mscorlib.dll.
'The hashed strings are calculated using calls to encryption methods built into mscorlib.dll.  The calculated hash strings are the same values as those calculated with jsSHA,
'a Javascript SHA implementation (see https://caligatio.github.io/jsSHA/ for an online calculator and the jsSHA code).
'The mscorlib.dll is intended for .Net Framework managed applications, but the stackoverflow.com post showed how it could be used with MS Access VBA.  This technique is
'not documented anywhere in MS Access documentation that I could find, so the stackoverflow.com post was very helpful in this regard.
'Requires a reference to mscorlib 4.0 64-bit
Dim Text As Object
Dim SHA512 As Object
Dim SHA256 As Object

Set Text = CreateObject("System.Text.UTF8Encoding")

Set SHA512 = CreateObject("System.Security.Cryptography.SHA512Managed")
Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")

Debug.Print B64Xml(SHA512.ComputeHash_2((Text.GetBytes_4("mypassword"))))
Debug.Print StrB64zXml(SHA512.ComputeHash_2((Text.GetBytes_4("mypassword"))))
End Sub

Private Sub WTst()
Dim X
Set X = CreateObject("System.Collections.ArrayList")
X.Add 1
Dim J%
For J = 1 To 1000
    X.Add J
Next
Dim I
For Each I In X
    Debug.Print I
Next
Stop
End Sub
