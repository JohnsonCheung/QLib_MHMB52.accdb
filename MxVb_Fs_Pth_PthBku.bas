Attribute VB_Name = "MxVb_Fs_Pth_PthBku"
#If Doc Then
'Bku:Cml   #Backup# used as verb
'Bu:Cml    #Backup# used as adjustive
'Fsi:Cml   #FileSystem-Item# Ffn or Pth
'P:Cml     #CurPj#
'C:Cml     #Cur#
'Ffn:Cml   #Full-File-Name#
'Pth:Cml   #Path#   A string of full path, optionally having path-separator as last char, which is preferred.
#End If
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Pth_OpBku."
Function BkuFfn$(Ffn, Optional Msg$ = "Bku")
Const CSub$ = CMod & "BkuFfn"
Dim Tmpn$:       Tmpn = Tmpn
Dim TarFfn$:   TarFfn = FfnBku(Ffn)
Dim MsgFfn$:   MsgFfn = Pth(TarFfn) & "Msg.txt"
Dim MsgiFfn$: MsgiFfn = PthFfnPar(TarFfn) & "MsgIdx.txt"
Dim Msgi$:       Msgi = "#" & Tmpn & vbTab & Msg & vbCrLf
:                       CpyFfn Ffn, TarFfn      ' <==
:                       WrtStr Msgi, MsgFfn     ' <==
:                       AppStr Msgi, MsgiFfn    ' <==
:                       BkuFfn = TarFfn
:                       Inf CSub, "File is Backuped", "As-file", TarFfn
End Function

Function PthBku$(Ffn): PthBku = PthAddFdrEns(PthAss(Ffn), ".backup"):      End Function ' :Pth #Backup-Path# ! The path used backuping @Ffn
Function FfnBku$(Ffn): FfnBku = PthAddFdrEns(PthBku(Ffn), Tmpn) & Fn(Ffn): End Function

Function FfnBkuLasPC$():  FfnBkuLasPC = FfnBkuLas(CPjf):      End Function
Function FfnBkuLas$(Ffn):   FfnBkuLas = EleMax(FfnyBku(Ffn)): End Function
Private Function WFdry(Fdry$()) As String()
Stop
End Function

Function IsFdrBku(Fdr) As Boolean
Stop
End Function
Function FfnyBku(Ffn) As String()
Dim P$: P = PthBku(Ffn)
Dim Fdr$(): Fdr = Fdry(P)
Dim Fn1$: Fn1 = Fn(Ffn)
Dim F: For Each F In Itr(Fdr)
    Dim FfnI$: FfnI = P & F & "\" & Fn1
    If HasFfn(FfnI) Then
        PushI FfnyBku, FfnI
    End If
Next
End Function
