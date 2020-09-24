Attribute VB_Name = "MxIde_Pj_Rf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Rf."
Function Rfln$(R As VbIde.Reference)
With R
Rfln = JnSpcAp(.Name, .Guid, .Major, .Minor, .FullPath)
End With
End Function

Private Sub B_RffLn()
Dim S$(): S = SrcRfP(CPj)
Dim I: For Each I In S
    Debug.Print RffLn(I)
Next
End Sub
Function RffLn(Rfln)
Const CSub$ = CMod & "RffLn"
Dim P%: P = InStr(Replace(Rfln, " ", "-", Count:=3), " ")
If P = 0 Then Thw CSub, "Invalid Rfln", "Rfln", Rfln
RffLn = Mid(Rfln, P + 1)
End Function

Function HasRfn(P As VBProject, Rfn) As Boolean: HasRfn = HasItn(P.References, Rfn):               End Function
Function NoRfn(P As VBProject, Rfn) As Boolean:   NoRfn = Not HasRfn(P, Rfn):                      End Function
Function HasRff(P As VBProject, Rff) As Boolean: HasRff = HasItppv(P.References, "FullPath", Rff): End Function

Function FtRfPC$():               FtRfPC = FtRfP(CPj):                   End Function
Function FtRfP$(P As VBProject):   FtRfP = FtRfPth(PthSrcP(P)):          End Function
Function FtRfPth$(PthSrc$):      FtRfPth = PthEnsSfx(PthSrc) & "Rf.txt": End Function

Function SrcRfPth(PthSrc$) As String(): SrcRfPth = LyFt(FtRfPth(PthSrc)): End Function
Function SrcRfPC() As String():          SrcRfPC = SrcRfP(CPj):           End Function
Function SrcRfP(P As VBProject) As String()
Dim R As VbIde.Reference: For Each R In P.References
    PushI SrcRfP, Rfln(R)
Next
End Function

Function CvRf(A) As VbIde.Reference: Set CvRf = A: End Function

Function HasRfGuid(P As VBProject, RfGuid): HasRfGuid = HasItppv(P.References, "GUID", RfGuid): End Function

Function RffyP(P As VBProject) As String(): RffyP = SyItp(P.References, "FullPath"): End Function

Function Rf(Rfn) As VbIde.Reference:                   Set Rf = RfP(CPj, Rfn):       End Function
Function RfP(P As VBProject, Rfn) As VbIde.Reference: Set RfP = CPj.References(Rfn): End Function
Function RfnyPC() As String():                         RfnyPC = RfnyP(CPj):          End Function
Function RfnyP(P As VBProject) As String():             RfnyP = Itn(P.References):   End Function
Sub DmpRf():                                                    DmpDrs DrsTRfPC:     End Sub

Function SrcLibRfUsr() As String()
Erase XX
X "MVb"
X "MIde  MVb MXls MAcs"
X "MXls  MVb"
X "MDao  MVb MDta"
X "MAdo  MVb"
X "MAdoX MVb"
X "MApp  MVb"
X "MDta  MVb"
X "MTp   MVb"
X "MSql  MVb"
X "AStkShpCst MVb MXls MAcs"
X "MAcs  MVb MXls"
SrcLibRfUsr = XX
Erase XX
End Function

Function SrcLibRfStd() As String()
Erase XX
X "QVb   Scripting VBScript_RegExp_55 DAO VBIDE Office"
X "QIde  Scripting VBIDE Excel"
X "QXls  Scripting Office Excel"
X "QDao  Scripting DAO"
X "QAdo  Scripting ADODB"
X "QAdoX Scripting ADOX"
X "QApp  Scripting"
X "QDta  Scripting"
X "Qtp   Scripting"
X "QSql  Scripting"
X "QAcs  Scripting Office Access"
X "QMHStkShpCst Scripting"
SrcLibRfStd = XX
End Function

Function SrcRfStd() As String()
Erase XX
X "VBA                {000204EF-0000-0000-C000-000000000046} 4  2 C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
X "Access             {4AFfC9A0-5F99-101B-AF4E-00AA003F0F07} 9  0 C:\Program Files (x86)\Microsoft Office\Root\Office16\MSACC.OLB"
X "stdole             {00020430-0000-0000-C000-000000000046} 2  0 C:\Windows\SysWOW64\stdole2.tlb"
X "Excel              {00020813-0000-0000-C000-000000000046} 1  9 C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE"
X "Scripting          {420B2830-E718-11CF-893D-00A0C9054228} 1  0 C:\Windows\SysWOW64\scrrun.dll"
X "DAO                {4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28} 12 0 C:\Program Files (x86)\Common Files\Microsoft Shared\OFfICE16\ACEDAO.DLL"
X "Office             {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52} 2  8 C:\Program Files (x86)\Common Files\Microsoft Shared\OFfICE16\MSO.DLL"
X "ADODB              {B691E011-1797-432E-907A-4D8C69339129} 6  1 C:\Program Files (x86)\Common Files\System\ado\msado15.dll"
X "ADOX               {00000600-0000-0010-8000-00AA006D2EA4} 6  0 C:\Program Files (x86)\Common Files\System\ado\msadox.dll"
X "VBScript_RegExp_55 {3F4DACA7-160D-11D2-A8E9-00104B365C9F} 5  5 C:\Windows\SysWOW64\vbscript.dll"
X "VBIDE              {0002E157-0000-0000-C000-000000000046} 5  3 C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
SrcRfStd = XX
Erase XX
End Function
Function DrsTRfStd() As Drs: DrsTRfStd = DrsTmy4R("Rfn Guid Maj Mnr Ffn", SrcRfStd): End Function
Sub BrwTRfStd():                         BrwDrs DrsTRfStd:                           End Sub
