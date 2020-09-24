Attribute VB_Name = "MxIde_Pj_Rf_Dic"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Rf_Dic."

Function DiRfnToRffPC() As Dictionary
Set DiRfnToRffPC = DiRfnToRffP(CPj)
End Function

Function DiRfnToRffP(P As VBProject) As Dictionary
Dim R As VbIde.Reference: For Each R In P.References
    DiRfnToRffP.Add R.Name, R.FullPath
Next
End Function

Function DiRfnToRffStd() As Dictionary
Static O As Dictionary
If Not IsNothing(O) Then GoTo X
Set O = New Dictionary
PushKvln O, "Excel              C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE"
PushKvln O, "stdole             C:\Windows\SysWOW64\stdole2.tlb"
PushKvln O, "Office             C:\Program Files (x86)\Common Files\Microsoft Shared\OFfICE16\MSO.DLL"
PushKvln O, "Access             C:\Program Files (x86)\Microsoft Office\Root\Office16\MSACC.OLB"
PushKvln O, "Scripting          C:\Windows\SysWOW64\scrrun.dll"
PushKvln O, "DAO                C:\Program Files (x86)\Common Files\Microsoft Shared\OFfICE16\ACEDAO.DLL"
PushKvln O, "ADODB              C:\Program Files (x86)\Common Files\System\ado\msado15.dll"
PushKvln O, "ADOX               C:\Program Files (x86)\Common Files\System\ado\msadox.dll"
PushKvln O, "VBScript_RegExp_55 C:\Windows\SysWOW64\vbscript.dll\3"
PushKvln O, "VBIDE              C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
PushKvln O, "MSForms            C:\WINDOWS\SysWOW64\FM20.DLL"
X: Set DiRfnToRffStd = O
End Function
