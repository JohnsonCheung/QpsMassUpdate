Attribute VB_Name = "Vb_Var"
Option Explicit

Function IsDte(S) As Boolean
On Error GoTo X
Dim A As Date
    A = S
    IsDte = True
    Exit Function
X:
End Function

Function Min(A, B, ParamArray Ap())
Dim O
    If A > B Then
        O = B
    Else
        O = A
    End If
Dim Av()
    Av = Ap
Dim J%
For J = 0 To UB(Av)
    If O > Av(J) Then
        O = Av(J)
    End If
Next
Min = O
End Function

Function Var_IsNothing(V) As Boolean
Var_IsNothing = TypeName(V) = "Nothing"
End Function

Private Sub Min__Tst()
Debug.Assert Min(1, 2, 4, 10, -1) = -1
End Sub
