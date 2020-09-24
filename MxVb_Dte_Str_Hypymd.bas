Attribute VB_Name = "MxVb_Dte_Str_Hypymd"
Option Compare Text
Option Explicit
Function Hypymd$(Y As Byte, M As Byte, D As Byte):    Hypymd = Hypym(Y, M) & "-" & Format(D, "00"): End Function
Function Hypym$(Y As Byte, M As Byte):                 Hypym = 2000 + Y & "-" & Format(M, "00"):    End Function
Function HypymYm$(A As Ym):                          HypymYm = Hypym(A.Y, A.M):                     End Function
Function HypymdYmd$(A As Ymd):                     HypymdYmd = Hypymd(A.Y, A.M, A.D):               End Function
Function HypymdNow$():                             HypymdNow = HypymdDte(Now):                      End Function
Function HypymdDte$(D As Date):                    HypymdDte = HypymdYmd(YmdDte(D)):                End Function
