Attribute VB_Name = "MxIde_Src_Msrc_WrtMsrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Msrc_MsrcIO."
Private Fso As New FileSystemObject
Private Const SepPth$ = "\"
Private Enum eCas: eCasIgn: eCasSen: End Enum 'Deriving(Str Val Txt)
Sub WrtMsrcy(S() As Nly)
ClrPthMsrc
Dim J&: For J = 0 To UbNly(S)
    With S(J)
        Dim Mdn$: Mdn = .Nm
        Dim Oldy$(): Oldy = SrcMdn(.Nm)
        Dim Newy$(): Newy = .Ly
    End With
    WrtAy Oldy, WWFtMdn(Mdn, "Old")  '<==
    WrtAy Newy, WWFtMdn(Mdn, "New") '<==
Next
Debug.Print "WrtMsrcy: NMsrcNew="; NMsrcNew
End Sub
Sub BrwPthMsrc(): BrwPth WWPthMsrc: End Sub
Private Function WWPthMsrc$()
Static X$: If X = "" Then X = PthAddFdrEns(PthAssPC, ".Msrc")
WWPthMsrc = X
End Function
Sub ClrPthMsrc():                                     ClrPth WWPthMsrc:                         End Sub
Function NMsrcNew%():                      NMsrcNew = Si(WWFfnyNew):                            End Function
Private Function WWFfnyNew() As String(): WWFfnyNew = Ffny(WWPthMsrc, "*(new).txt"):            End Function
Private Function WWFtMdn$(Mdn$, NewOld$):   WWFtMdn = WWPthMsrc & Mdn & "(" & NewOld & ").txt": End Function

Sub VcPthMsrc(): VcPth WWPthMsrc: End Sub
