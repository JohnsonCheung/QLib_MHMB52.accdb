Attribute VB_Name = "MxDao_Db_Schm_Ud"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Schm_Ud."
'== Sms SchmSrc =================================================
Type SmsTbl:       Lno As Integer: Tbn As String:  Fny() As String: FnySk() As String:            End Type 'Deriving(Ctor Ay)
Type SmsEleFld:    Lno As Integer: Elen As String: FldLiky() As String:                           End Type 'Deriving(Ctor Ay)
Type SmsEle:       Lno As Integer: Elen As String: EleStr As String:                              End Type 'Deriving(Ctor Ay)
Type SmsKey:       Lno As Integer: Tbn As String:  Keyn As String: IsUniq As Boolean: Fny() As String: End Type 'Deriving(Ctor Ay)
Type SmsTblDes:    Lno As Integer: Tbn As String:  Des As String:                                 End Type 'Deriving(Ctor Ay)
Type SmsTblFldDes: Lno As Integer: Tbn As String:  Fldn As String: Des As String:                 End Type 'Deriving(Ctor Ay)
Type SmsFldDes:    Lno As Integer: Fldn As String: Des As String: End Type 'Deriving(Ctor Ay)
Type SchmSrc
    Tbl() As SmsTbl
    EleFld() As SmsEleFld
    Ele() As SmsEle
    PvTDes() As SmsTblDes
    TblFldDes() As SmsTblFldDes
    PvFDesC() As SmsFldDes
    Key() As SmsKey
End Type
Function SmsTblAdd(A As SmsTbl, B As SmsTbl) As SmsTbl(): PushSmsTbl SmsTblAdd, A: PushSmsTbl SmsTblAdd, B: End Function
Sub PushSmsTbly(O() As SmsTbl, A() As SmsTbl): Dim J&: For J = 0 To UbSmsTbl(A): PushSmsTbl O, A(J): Next: End Sub
Sub PushSmsTbl(O() As SmsTbl, M As SmsTbl): Dim N&: N = SiSmsTbl(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiSmsTbl&(A() As SmsTbl): On Error Resume Next: SiSmsTbl = UBound(A) + 1: End Function
Function UbSmsTbl&(A() As SmsTbl): UbSmsTbl = SiSmsTbl(A) - 1: End Function
Function SmsTbl(Lno, Tbn, Fny$(), FnySk$()) As SmsTbl
With SmsTbl
    .Lno = Lno
    .Tbn = Tbn
    .Fny = Fny
    .FnySk = FnySk
End With
End Function
Function SmsEleFldAdd(A As SmsEleFld, B As SmsEleFld) As SmsEleFld(): PushSmsEleFld SmsEleFldAdd, A: PushSmsEleFld SmsEleFldAdd, B: End Function
Sub PushSmsEleFldy(O() As SmsEleFld, A() As SmsEleFld): Dim J&: For J = 0 To UbSmsEleFld(A): PushSmsEleFld O, A(J): Next: End Sub
Sub PushSmsEleFld(O() As SmsEleFld, M As SmsEleFld): Dim N&: N = SiSmsEleFld(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiSmsEleFld&(A() As SmsEleFld): On Error Resume Next: SiSmsEleFld = UBound(A) + 1: End Function
Function UbSmsEleFld&(A() As SmsEleFld): UbSmsEleFld = SiSmsEleFld(A) - 1: End Function
Function SmsEleFld(Lno, Elen, FldLiky$()) As SmsEleFld
With SmsEleFld
    .Lno = Lno
    .Elen = Elen
    .FldLiky = FldLiky
End With
End Function
Function SmsEleAdd(A As SmsEle, B As SmsEle) As SmsEle(): PushSmsEle SmsEleAdd, A: PushSmsEle SmsEleAdd, B: End Function
Sub PushSmsEley(O() As SmsEle, A() As SmsEle): Dim J&: For J = 0 To UbSmsEle(A): PushSmsEle O, A(J): Next: End Sub
Sub PushSmsEle(O() As SmsEle, M As SmsEle): Dim N&: N = SiSmsEle(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiSmsEle&(A() As SmsEle): On Error Resume Next: SiSmsEle = UBound(A) + 1: End Function
Function UbSmsEle&(A() As SmsEle): UbSmsEle = SiSmsEle(A) - 1: End Function
Function SmsEle(Lno, Elen, EleStr) As SmsEle
With SmsEle
    .Lno = Lno
    .Elen = Elen
    .EleStr = EleStr
End With
End Function
Function SmsKeyAdd(A As SmsKey, B As SmsKey) As SmsKey(): PushSmsKey SmsKeyAdd, A: PushSmsKey SmsKeyAdd, B: End Function
Sub PushSmsKeyy(O() As SmsKey, A() As SmsKey): Dim J&: For J = 0 To UbSmsKey(A): PushSmsKey O, A(J): Next: End Sub
Sub PushSmsKey(O() As SmsKey, M As SmsKey): Dim N&: N = SiSmsKey(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiSmsKey&(A() As SmsKey): On Error Resume Next: SiSmsKey = UBound(A) + 1: End Function
Function UbSmsKey&(A() As SmsKey): UbSmsKey = SiSmsKey(A) - 1: End Function
Function SmsKey(Lno, Tbn, Keyn, IsUniq, Fny$()) As SmsKey
With SmsKey
    .Lno = Lno
    .Tbn = Tbn
    .Keyn = Keyn
    .IsUniq = IsUniq
    .Fny = Fny
End With
End Function
Function SmsTblDesAdd(A As SmsTblDes, B As SmsTblDes) As SmsTblDes(): PushSmsTblDes SmsTblDesAdd, A: PushSmsTblDes SmsTblDesAdd, B: End Function
Sub PushSmsTblDesy(O() As SmsTblDes, A() As SmsTblDes): Dim J&: For J = 0 To UbSmsTblDes(A): PushSmsTblDes O, A(J): Next: End Sub
Sub PushSmsTblDes(O() As SmsTblDes, M As SmsTblDes): Dim N&: N = SiSmsTblDes(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiSmsTblDes&(A() As SmsTblDes): On Error Resume Next: SiSmsTblDes = UBound(A) + 1: End Function
Function UbSmsTblDes&(A() As SmsTblDes): UbSmsTblDes = SiSmsTblDes(A) - 1: End Function
Function SmsTblDes(Lno, Tbn, Des) As SmsTblDes
With SmsTblDes
    .Lno = Lno
    .Tbn = Tbn
    .Des = Des
End With
End Function
Function SmsTblFldDesAdd(A As SmsTblFldDes, B As SmsTblFldDes) As SmsTblFldDes(): PushSmsTblFldDes SmsTblFldDesAdd, A: PushSmsTblFldDes SmsTblFldDesAdd, B: End Function
Sub PushSmsTblFldDesy(O() As SmsTblFldDes, A() As SmsTblFldDes): Dim J&: For J = 0 To UbSmsTblFldDes(A): PushSmsTblFldDes O, A(J): Next: End Sub
Sub PushSmsTblFldDes(O() As SmsTblFldDes, M As SmsTblFldDes): Dim N&: N = SiSmsTblFldDes(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiSmsTblFldDes&(A() As SmsTblFldDes): On Error Resume Next: SiSmsTblFldDes = UBound(A) + 1: End Function
Function UbSmsTblFldDes&(A() As SmsTblFldDes): UbSmsTblFldDes = SiSmsTblFldDes(A) - 1: End Function
Function SmsTblFldDes(Lno, Tbn, Fldn, Des) As SmsTblFldDes
With SmsTblFldDes
    .Lno = Lno
    .Tbn = Tbn
    .Fldn = Fldn
    .Des = Des
End With
End Function
Function SmsFldDesAdd(A As SmsFldDes, B As SmsFldDes) As SmsFldDes(): PushSmsFldDes SmsFldDesAdd, A: PushSmsFldDes SmsFldDesAdd, B: End Function
Sub PushSmsFldDesy(O() As SmsFldDes, A() As SmsFldDes): Dim J&: For J = 0 To UbSmsFldDes(A): PushSmsFldDes O, A(J): Next: End Sub
Sub PushSmsFldDes(O() As SmsFldDes, M As SmsFldDes): Dim N&: N = SiSmsFldDes(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiSmsFldDes&(A() As SmsFldDes): On Error Resume Next: SiSmsFldDes = UBound(A) + 1: End Function
Function UbSmsFldDes&(A() As SmsFldDes): UbSmsFldDes = SiSmsFldDes(A) - 1: End Function
Function SmsFldDes(Lno, Fldn, Des) As SmsFldDes
With SmsFldDes
    .Lno = Lno
    .Fldn = Fldn
    .Des = Des
End With
End Function
