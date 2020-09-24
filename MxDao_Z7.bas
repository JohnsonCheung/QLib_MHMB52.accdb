Attribute VB_Name = "MxDao_Z7"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_A_Z7."

Sub UnzipPth(Dbz7 As Database, FfnZip, PthTo)
If NoPth(PthPar(PthTo)) Then SetMainMsg "To Path not found: " & PthPar(PthTo): Exit Sub
If NoFfn(FfnZip) Then SetMainMsg "Zip file not found: " & FfnZip: Exit Sub
If IsCfmAndClrPthR(PthTo) Then Exit Sub

Dim P$:             P = PthAddFdrEns(Pth(Dbz7.Name), "ZipPthWrking")
Dim Fexe$:       Fexe = P & "z7.exe"
Dim Fcmd$:       Fcmd = P & "Unzip.Cmd"
Dim FcmdStr$: ' FcmdStr = WCxt(P, FfnZip, PthTo)
Dim ShellStr$: ShellStr = FmtQQ("Cmd.Exe /C ""?""", Fcmd)

                        WExpz7 Dbz7
                        WrtStr FcmdStr, Fcmd
                        Shell Fcmd, vbMaximizedFocus
End Sub

Private Function WCxtUnzip$(FfnZip, PthTo)
Dim O$()
Dim W$: W = WPthz7
Dim Z$: Z = W & Fn(FfnZip)
Dim T$: T = PthPar(PthTo)       ' Target-Path should the parent of @PthTo
Push O, FmtQQ("Cd ""?""", T)
Push O, ""
Push O, FmtQQ("Copy  ""?"" ""?""", FfnZip, W)
Push O, ""
Push O, FmtQQ("""?z7.exe"" x -r ""?""", W, Z)
Push O, ""
Push O, FmtQQ("Del ""?""", Z)
Push O, "Pause"
WCxtUnzip = JnCrLf(O)
End Function

Private Sub B_UnzipPth()
Dim B$: B = ResPthB
Dim A$: A = ResPthA
CrtResPthA
PthEns B
ZipPth CurrentDb, A, B
Stop ' Wait the Dos to zip
Stop 'UnzipPth CurrentDb, Z, A
BrwPth A '<== A should be restored
End Sub
Sub BrwPthZ7(): BrwPth WPthz7: End Sub
Private Sub B_WExp7z()
Dim Dbz7 As Database
Set Dbz7 = CDb
WExpz7 CDb
MsgBox HasFfn(WFexez7)
End Sub
Private Function WPthz7$()
Static P$: If P = "" Then P = PthAddFdrEns(PthTmpRoot, "Z7")
WPthz7 = P
End Function
Private Function WFexez7$():       WFexez7 = WPthz7 & "z7.exe":                                                         End Function
Private Function WFoup$(PthFm):      WFoup = Fdr(PthFm) & "(" & Format(Now, "YYYY-MM-DD HH-MM") & " " & CUsr & ").zip": End Function
Private Function WFcmdZip$():     WFcmdZip = WPthz7 & "ZipPth.cmd":                                                     End Function
Private Function WFcmdUnzip$(): WFcmdUnzip = WPthz7 & "UnzipPth.cmd":                                                   End Function
Sub ZipPth(Dbz7 As Database, PthFm, PthTo)
WChkPmEr PthFm, PthTo
WExpz7 Dbz7
WCrtFcmdZip PthFm, PthTo
Shell WStrShellZip, vbMaximizedFocus
End Sub
Private Function WStrShellZip$():           WStrShellZip = WStrShell(WFcmdZip):                                       End Function
Private Function WStrShellUnzip$():       WStrShellUnzip = WStrShell(WFcmdUnzip):                                     End Function
Private Function WStrShell$(Fcmd$):            WStrShell = FmtQQ("Cmd.exe /C ""?""", Fcmd):                           End Function
Private Sub WCrtFcmdZip(PthFm, PthTo):                     WrtStr WCxtZip(PthFm, PthTo), WFcmdZip, OvrWrt:=True:      End Sub
Private Sub WCrtFcmdUnzip(FfnZip, PthTo):                  WrtStr WCxtUnzip(FfnZip, PthTo), WFcmdUnzip, OvrWrt:=True: End Sub
Private Function WCxtZip$(PthFm, PthTo)
':Cxt: :Lines ! #File-Context#
Dim O$()
Dim Foup$: Foup = WFoup(PthFm)
Push O, FmtQQ("Cd ""?""", WPthz7)
Push O, FmtQQ("z7 a ""?"" ""?""", Foup, PthFm)
Push O, FmtQQ("move ""?"" ""?""", Foup, PthTo)
Push O, "Pause"
WCxtZip = JnCrLf(O)
End Function

Private Sub WExpz7(Dbz7 As Database)
Dim Fexe$: Fexe = WFexez7
If HasFfn(Fexe) Then Exit Sub
Dim R As Dao.Recordset: Set R = Dbz7.TableDefs("QLib_7z").OpenRecordset
Dim R2 As Dao.Recordset2: Set R2 = R.Fields("7z").Value
Dim F2 As Dao.Field2: Set F2 = R2.Fields("FileData")
Dim NoExt$: NoExt = Ffnn(Fexe)
DltFfnIf NoExt
F2.SaveToFile NoExt
Name NoExt As Fexe
End Sub
Private Sub WChkPmEr(PthFm, PthTo)
Dim O$()
If NoPth(PthTo) Then PushI O, "To Path not found: " & PthTo
If NoPth(PthFm) Then PushI O, "From Path not found: " & PthFm
End Sub

Private Sub B_ZipPth(): ZipPth CDb, PthTmp, PthTmpInst: End Sub
