Attribute VB_Name = "MxIde_Lis_LisMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Lis_LisMd."

Sub BrwMd(Optional PatnssAndMd$):                        BrwAy WFmtMd(PatnssAndMd):               End Sub
Sub BrwMdSen(Optional PatnssAndMd$):                     BrwAy WFmtMd(PatnssAndMd, eCasSen):      End Sub
Sub BrwMdf(Mdfssub$):                                    BrwAy WFmtMdf(Mdfssub):                  End Sub
Sub BrwMdfSen(Mdfssub$):                                 BrwAy WFmtMdf(Mdfssub, eCasSen):         End Sub
Sub LisMd(Optional PatnssAndMd$, Optional Top& = 50):    DmpAy WFmtMd(PatnssAndMd, , Top):        End Sub
Sub LisMdSen(Optional PatnssAndMd$, Optional Top& = 50): DmpAy WFmtMd(PatnssAndMd, eCasSen, Top): End Sub
Sub LisMdf(Mdfssub$, Optional Top& = 50):                DmpAy WFmtMdf(Mdfssub, , Top):           End Sub
Sub LisMdfSen(Mdfssub$, Optional Top& = 50):             DmpAy WFmtMdf(Mdfssub, eCasSen, Top):    End Sub
Sub VcMd(Optional PatnssAndMd$):                         VcAy WFmtMd(PatnssAndMd):                End Sub
Sub VcMdSen(Optional PatnssAndMd$):                      VcAy WFmtMd(PatnssAndMd, eCasSen):       End Sub
Sub VcMdf(Mdfssub$):                                     VcAy WFmtMdf(Mdfssub):                   End Sub
Sub VcMdfSen(Mdfssub$):                                  VcAy WFmtMdf(Mdfssub, eCasSen):          End Sub
Private Function WFmtMd(PatnssAndMd$, Optional C As eCas, Optional Top&) As String()
Dim Ny$(): Ny = MdnyPC(PatnssAndMd)
WFmtMd = AwFstNSrt(Ny, Top)
End Function
Private Function WFmtMdf(Mdfssub$, Optional C As eCas, Optional Top&) As String()
Dim Mdny$(): Mdny = MdnyMdfssub(Mdfssub, C)
WFmtMdf = AwFstNSrt(Mdny, Top)
End Function
