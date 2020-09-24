Attribute VB_Name = "MxIde_Mdn_Imdn"
Option Compare Text
Option Explicit
Function MdnyIntlPC() As String(): MdnyIntlPC = WAwMdnIsIntl(MdnyPC, IsIntl:=True): End Function ' Internal Mdn
Function MdnyExtlPC() As String(): MdnyExtlPC = WAwMdnIsIntl(MdnyPC, IsIntl:=False): End Function ' External Mdn

Private Function WAwMdnIsIntl(Mdny$(), IsIntl As Boolean) As String()
Dim Mdn: For Each Mdn In Itr(Mdny)
    If WIsMdnSel(Mdn, IsIntl) Then PushI WAwMdnIsIntl, Mdn
Next
End Function
Private Function WIsMdnSel(Mdn, IsIntl As Boolean) As Boolean
Dim HasIntl As Boolean
Select Case True
Case HasSsub(Mdn, "_Intl_"), HasSsub(Mdn, "_Tool_"): HasIntl = True
End Select
If IsIntl Then
    WIsMdnSel = HasIntl
Else
    WIsMdnSel = Not HasIntl
End If
End Function
