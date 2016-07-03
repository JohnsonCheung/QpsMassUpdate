Attribute VB_Name = "Vb_Fmt"
Option Explicit

Function Fmt$(FmtStr$, ParamArray Ap())
Dim Av()
Av = Ap
Fmt = Fmt_Av(FmtStr, Av)
End Function

Function Fmt_Av$(FmtStr$, Av())
Dim I, O$, J%, A$
O = FmtStr
For Each I In Av
    A = Quote(J, "{}"): J = J + 1
    O = Replace(O, A, I)
Next
Fmt_Av = O
End Function

Function Fmt_ErDes$(ErMsg$, Av)
Dim A$()
    A = Str_MacroAy(ErMsg)
If Sz(A) <> Sz(Av) Then
    Stop
    Fmt_ErDes = ErMsg
    Exit Function
End If
Dim O$()
    Push O, ErMsg
    Push O, ""
    Push O, ""
    Dim J%
    For J = 0 To UB(A)
        If IsArray(Av(J)) Then
            Push O, A(J) & " = [" & Ay_ToStr(Av(J)) & "]"
        Else
            Push O, A(J) & " = [" & Av(J) & "]"
        End If
        Push O, ""
    Next
Fmt_ErDes = Join(O, vbCrLf)
End Function

Function Fmt_QQ$(QQ$, ParamArray Ap())
Dim Av()
Av = Ap
Fmt_QQ = Fmt_QQAv(QQ, Av)
End Function

Function Fmt_QQAv$(QQ$, Av())
Dim I, O$
O = QQ
For Each I In Av
    O = Replace(O, "?", CStr(I), Count:=1)
Next
Fmt_QQAv = O
End Function

Private Function Str_MacroAy(S$) As String()
Dim A$
    A = S
Dim O$()
    Dim P%
        P = InStr(A, "{")
    Do While P > 0
        A = Mid(A, P)
        P = InStr(A, "}")
        If P = 0 Then Er "{S} only has open-angel-bracket, but no closing", S
        Push O, Left(A, P)
        A = Mid(A, P + 1)
        P = InStr(A, "{")
    Loop
Str_MacroAy = O
End Function
