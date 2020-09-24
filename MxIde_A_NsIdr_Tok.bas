Attribute VB_Name = "MxIde_A_NsIdr_Tok"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_A_NsIdr_Tok."
Private Const WTySS$ = "Boolean Integer Double Single Currency Boolean Byte"
Public Const SsVbKw$ = WTySS & "As Compare Const Dim Do Each Else End Empty Exit Explicit False For Function Get If In Loop Me New Next Not" & _
" Option Optional Private Property Set Static Sub Text Then To True Variant Wend While"
Public Enum eTok: eTokNone: eTokUnknown
: eTokRmk: eTokLitNum: eTokLitStr: eTokLitBool: eTokLitDte
: eTokBktOpn: eTokBktCls: eTokDot
: eTokOpLE: eTokOpLT: eTokOpGE: eTokOpGT: eTokOpEQ
: eTokOpAdd: eTokOpDiv: eTokOpMulti: eTokOpMinus
: eTokKw: eTokIdr
End Enum
Type Tok: Ty As eTok: Val As String: End Type
Type TokOpt: Som As Boolean: Tok As Tok: End Type
Function SyVbKw() As String()
Static X$()
If Si(X) = 0 Then
    X = SySs(SsVbKw)
End If
SyVbKw = X
End Function
Sub PushTok(O() As Tok, M As Tok)
Dim N&: Stop ' N = Tok_Si(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Function TokyL(Ln) As Tok()
Dim L$: L = Ln
While L <> ""
    With ShfTok(L)
       Stop ' If .Som Then
            'PushTok , TokyL, .Tok
        'Else
            Exit Function
        'End If
    End With
Wend
End Function
Function ShfTok(OContln$) As Tok
Stop 'OContln = LTrim(OContln): If L = "" Then OContln = "": Exit Function
Dim O As TokOpt
With O.Tok
    Dim C1$: C1 = ChrFst(OContln)
    Select Case True
    Case C1 = ":"
    Case C1 = """": Stop 'O = WShf_LitStr(OContln)
    Case C1 = "#": Stop 'O = WShf_LitDte(OContln)
    Case IsDig(C1): Stop 'O = WShf_LitNum(OContln)
    Case C1 = "'": .Ty = eTokRmk: .Val = Mid(OContln, 2): OContln = ""
    Case C1 = "(": .Ty = eTokBktOpn: OContln = Mid(OContln, 2)
    Case C1 = ")": .Ty = eTokBktCls: OContln = Mid(OContln, 2)
    Case C1 = ".": .Ty = eTokDot: OContln = Mid(OContln, 2)
    Case C1 = ",": Stop '.Ty = eTokCma: OContln = Mid(OContln, 2)
    Case C1 = "=": .Ty = eTokOpEQ: OContln = Mid(OContln, 2)
    Case C1 = "<"
        If Mid(OContln, 2, 1) = "=" Then
            .Ty = eTokOpLE: OContln = Mid(OContln, 3)
        Else
            .Ty = eTokOpLT: OContln = Mid(OContln, 2)
        End If
    Case C1 = "<"
        If Mid(OContln, 2, 1) = "=" Then
            .Ty = eTokOpLE: OContln = Mid(OContln, 3)
        Else
            .Ty = eTokOpLT: OContln = Mid(OContln, 2)
        End If
    Case C1 = ">"
        If Mid(OContln, 2, 1) = "=" Then
            .Ty = eTokOpGE: OContln = Mid(OContln, 3)
        Else
            .Ty = eTokOpGT: OContln = Mid(OContln, 2)
        End If
    Case Else: .Ty = eTokUnknown: .Val = OContln: OContln = ""
    End Select
End With
End Function
Private Function WShf_LitStr$(OLn$)

End Function
Private Function WShf_LitDte$(OLn$)

End Function
Private Function WShf_LitNum$(OLn$)

End Function
Function LyToky(T() As Tok) As String()
Dim J&: Stop 'For J = 0 To Tok_UB(T)
    PushI LyToky, StrTok(T(J))
'Next
End Function
Function StrTok$(T As Tok): Stop ' StrTok = EnmsTok(T.Ty) & " " & T.Val: End Function

End Function
