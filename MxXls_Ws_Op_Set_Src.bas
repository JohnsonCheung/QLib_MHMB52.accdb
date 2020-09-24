Attribute VB_Name = "MxXls_Ws_Op_Set_Src"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_Op_Set_Src."
Private Sub B_AddCdlWsM()
Dim B As Workbook: Set B = WbNw
Dim S As Worksheet: Set S = B.Sheets("Sheet1")
AddCdlWsM S, CMd
Maxv S.Application
End Sub

Sub AddCdlWs(S As Worksheet, Cdl$):                 RplMd MdWs(S), Cdl:       End Sub
Sub AddCdlWsM(S As Worksheet, MdSrc As CodeModule): AddCdlWs S, SrclM(MdSrc): End Sub
Sub AddCdlWsMdn(S As Worksheet, Mdn$):              AddCdlWsM S, Md(Mdn):     End Sub
