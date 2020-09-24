Attribute VB_Name = "MxDao_Att_AttOp_IExpAttIf"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Att_Do_AttIExpIf."
Sub ExpAttIf(D As Database, Ffn$, Attn$, Optional Attfn$) ' If AttFfn is older -> Export; If AttFfn is older -> export
Const CSub$ = CMod & "ExpAttIf"
Select Case True
Case Not HasAttFnzTbAtt(D, Attn, Attfn):  Thw CSub, "Given Attn.Attf not found", "Attn Attf", Attn, Attfn
Case Not HasFfn(Ffn):                  GoSub Exp
Case ZIsAttNewer(Ffn, D, Attn, Attfn): DltFfn Ffn: GoSub Exp
End Select
Exit Sub
Exp:
    ExpAtt D, Attn, Attfn, Ffn
    Debug.Print FmtQQ("?: Ffn imported.  Attn[?] Attf[?] Ffn[?]", CSub, Attn, Attfn, Ffn)
    Return
End Sub
Sub ExpAttIfC(Ffn$, Attn$, Optional Attfn$): ExpAttIf CDb, Ffn, Attn, Attfn: End Sub
Sub ImpAttIf(D As Database, Ffn$, Attn$, Optional Attfn$) ' If AttFfn is older -> Export; If AttFfn is newer -> import
Const CSub$ = CMod & "ImpAttIf"
If Not HasFfn(Ffn) Then Exit Sub
Select Case True
Case Not HasFfn(Ffn):                  Inf CSub, FmtQQ("Ffn[?] not exist", Ffn)
Case Not HasAttFnzTbAtt(D, Attn, Attfn):  GoSub Import
Case ZIsAttOlder(Ffn, D, Attn, Attfn): GoSub Import
End Select
Exit Sub
Import:
    ImpAtt Ffn, D, Attn, Attfn
    Return
End Sub
Sub ImpAttIfC(Ffn$, Attn$, Optional Attfn$):                                               ExpAttIf CDb, Ffn, Attn, Attfn:          End Sub
Private Function ZIsAttOlder(Ffn$, D As Database, Attn$, Attfn$) As Boolean: ZIsAttOlder = TimAttf(D, Attn, Fn(Ffn)) < FfnTim(Ffn): End Function
Private Function ZIsAttNewer(Ffn$, D As Database, Attn$, Attfn$) As Boolean: ZIsAttNewer = TimAttf(D, Attn, Fn(Ffn)) > FfnTim(Ffn): End Function
