Attribute VB_Name = "Vb_Assert"
Option Explicit

Sub Assert_Eq(A1, A2, ErMsg$, ParamArray Ap())
If A1 <> A2 Then
    Dim Av()
        Av = Ap
    Er_ByAv ErMsg, Av
End If
End Sub

Sub Assert_NotEq(A1, A2, ErMsg$, ParamArray Ap())
If A1 = A2 Then
    Dim Av()
        Av = Ap
    Er_ByAv ErMsg, Av
End If
End Sub
