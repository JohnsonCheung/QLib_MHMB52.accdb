Attribute VB_Name = "MxIde_Lis_LisMthCac"
Option Compare Text
Option Explicit
Enum eEIntl: eEIntlBoth: eEItnlE: eEIntlI: End Enum
Enum eMthSrt: eMthSrtByMd: eMthSrtByMthn: End Enum
Public Const EnmqssEIntl$ = "eEIntl? Both E I"
Private Sub B_DmpMthCac(): LisMthCac , "CCml": End Sub

Sub BrwMthCac(Optional PatnssAndMd$, Optional PatnssAndMthn$, Optional PatnssAndMthln$, Optional S As eMthSrt, Optional M As eWhMdy, Optional EI As eEIntl): BrwAy WFmtMi3Mnl(PatnssAndMd, PatnssAndMthn, PatnssAndMthln, S, M, EI): End Sub
Sub LisMthCac(Optional PatnssAndMd$, Optional PatnssAndMthn$, Optional PatnssAndMthln$, Optional S As eMthSrt, Optional M As eWhMdy, Optional EI As eEIntl): DmpAy WFmtMi3Mnl(PatnssAndMd, PatnssAndMthn, PatnssAndMthln, S, M, EI): End Sub
Sub VcMthCac(Optional PatnssAndMd$, Optional PatnssAndMthn$, Optional PatnssAndMthln$, Optional S As eMthSrt, Optional M As eWhMdy, Optional EI As eEIntl, Optional C As eCas): VcAy WFmtMi3Mnl(PatnssAndMd, PatnssAndMthn, PatnssAndMthln, S, M, EI):  End Sub

Private Function WFmtMi3Mnl(PatnssAndMd$, PatnssAndMthn$, PatnssAndMthln$, S As eMthSrt, M As eWhMdy, EI As eEIntl) As String()
EnsFtcacPC
Dim RxMthn() As RegExp: RxMthn = Rxay(PatnssAndMthn)
Dim RxMthln() As RegExp: RxMthln = Rxay(PatnssAndMthln)
Dim Mdny1$(): Mdny1 = MdnyP(CPj, PatnssAndMd) 'Filter Mdn
Dim Mdny$(): Mdny = SySrtQ(MdnyWhEIntl(Mdny1, EI))   'Filter EItnl
Dim Dy1(): Dy1 = Dy4MnflzMdny(Mdny)              ' 4MdMthnMdyMthl_WhMdny
Dim Dy2(): Dy2 = DyWhRxayAnd(Dy1, 1, RxMthn)     ' Filter Mthn
Dim Dy3(): Dy3 = DyWhRxayAnd(Dy2, 3, RxMthln)    ' Filter Mthln
Dim Dy4(): Dy4 = DyWhMdy(Dy3, M)                ' FIlter Mdy
Dim Dy5(): Dy5 = DySel(Dy4, Inty(0, 1, 3))       ' Sel Md-Mthn-Mthln, skip the Mdy
Dim Dy6()
    If S = eMthSrtByMthn Then
        Stop
        Dy6 = DySrtCii(Dy5, 1)                      ' Srt by Mthn
    Else
        Dy6 = Dy5
    End If
Dim ODy(): ODy = DyAli(Dy6, W:=0)                ' Ali Dy
WFmtMi3Mnl = DyJn(ODy)
PushI WFmtMi3Mnl, "Cnt: " & Si(ODy)
End Function
Private Function Dy4MnflzMdny(Mdny$()) ' Read from FtCacMit8 return Dy4Mnfl (Md Mthn Mdy Mthln)
Dim N$() ' TsyMi8Cmntfbel
         '       12 4  7
    Dim Pth$: Pth = PthFtcacP(CPj)
    Dim Mdn: For Each Mdn In Itr(Mdny)
        Dim Ft$: Ft = Pth & Mdn & FnsfxFtcacMit8 & ".txt"
        PushIAy N, LyFtIf(Ft)
    Next
Dim Dy(): Dy = DyTsy(N)
Dy4MnflzMdny = DySel(Dy, Inty(1, 2, 4, 7)) ' 1 2 4 7 is Mnfl - Md-Mthn-Mdy-Mthln
End Function

Private Function DyWhMdy(Dy4Mnfl(), M As eWhMdy) As Variant()
If M = eWhMdyAll Then DyWhMdy = Dy4Mnfl: Exit Function
Dim Dr: For Each Dr In Itr(Dy4Mnfl)
    If HitMdy(Dr(2), M) Then    'Mnfl
                                '  2
        PushI DyWhMdy, Dr
    End If
Next
End Function
