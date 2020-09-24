Attribute VB_Name = "MxVb_Cnst"
Option Compare Text
Option Explicit
Public Const CMod$ = "?"
#If Doc Then
'Dfn-Rul:: To define some thing:
'          #1 'XXX::         Rx = "('[\w\-]+\:\:| [\w\-]+$"
'          #2 {Spc}XXX::
'          #3 only one
'Ref-Rul:: To ref
'          #1 {Spc}:XXX{Spc}  Rx = " :[\w\-]+ "
'          #2 {Spc}:XXX$      Rx = " :[\w\-]+$"
#End If
Public Const vbBktOpn$ = "("
Public Const vbSpc4$ = "    "
Public Const vbBktCls$ = "("
Public Const vbBktOpnBig$ = "{"
Public Const vbBktClsBig$ = "}"
Public Const vbBktOpnSq$ = "["
Public Const vbBktClsSq$ = "]"
Public Const vbSpc$ = " "
Public Const vbQuoDbl$ = """"
Public Const vbCma$ = ","
Public Const vbCmaCrLf$ = vbCma & vbCrLf
Public Const vbCmaSpc$ = vbCma & vbSpc
Public Const vbAscQuoDbl As Byte = 34
Public Const vbQuoSng$ = "'"
Public Const vbExm$ = "!"
Public Const vbHsh$ = "#"
Public Const vbCrLf2$ = vbCrLf & vbCrLf
Public Const vbQuoDbl2$ = vbQuoDbl & vbQuoDbl
Public Const vbQuoSng2$ = vbQuoSng & vbQuoSng
