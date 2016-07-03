Attribute VB_Name = "FmtChrCol"
Option Explicit

Sub FmtChrCol_AsTxt(Ws As Worksheet, TCno As TCno)
Dim NChr%
Dim ChrCno() As ChrCno
With TCno
    NChr = .NChr
    ChrCno = .Chr
End With
Dim J&
Dim R1&, R2
With Ws_ListObj_R1R2(Ws)
    R1 = .R1
    R2 = .R2
End With
For J = 0 To NChr - 1
    Ws_CRR(Ws, ChrCno(J).Cno, R1, R2).NumberFormat = "@"
Next
'--- Fmt to Txt is not OK, because after using Rge.Value, the value is still text
'--- Convert to Txt for numeric cell
For J = 0 To NChr - 1
    Dim Cno%
         Cno = ChrCno(J).Cno
    Dim VBar As Range
        Set VBar = Ws_CRR(Ws, Cno, R1, R2)
    
    Dim Sqv
        Sqv = VBar.Value
    
    Dim AnyNbr As Boolean
        AnyNbr = False
    If VBar.Count = 1 Then
        If VarType(Sqv) = vbDouble Then
            Sqv = CStr(Sqv)
            Debug.Print Fmt_QQ("FmtChrCol_AsTxt: Convert RC(1,?) to Txt[?]", Cno, Sqv)
        End If
    Else
        Dim I&
        For I = 1 To UBound(Sqv, 1)
            If VarType(Sqv(I, 1)) = vbDouble Then
                AnyNbr = True
                Sqv(I, 1) = CStr(Sqv(I, 1))
                Debug.Print Fmt_QQ("FmtChrCol_AsTxt: Convert RC(?,?) to Txt[?]", I, Cno, Sqv(I, 1))
            End If
        Next
    End If
    
    If AnyNbr Then
        VBar.Value = Sqv
    End If
Next
End Sub
