Attribute VB_Name = "MxXls_Ws_IdeobjWs"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Ws_Pj."
Function MdWs(S As Worksheet) As CodeModule:    Set MdWs = CmpWs(S).CodeModule:                        End Function
Function CmpWs(S As Worksheet) As VBComponent: Set CmpWs = ItoFstNm(PjWs(S).VBComponents, S.CodeName): End Function
Function PjWb(B As Workbook) As VBProject:      Set PjWb = B.VBProject:                                End Function
Function PjWs(S As Worksheet) As VBProject:     Set PjWs = WbWs(S).VBProject:                          End Function
Function PjRg(A As Range) As VBProject:         Set PjRg = WbRg(A).VBProject:                          End Function
