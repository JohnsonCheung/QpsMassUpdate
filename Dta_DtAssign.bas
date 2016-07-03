Attribute VB_Name = "Dta_DtAssign"
Option Explicit

Sub Dt_AssignRow(Dt As TDt, R&, CnoAy%(), ParamArray OAp())
Dim OAv()
    OAv = OAp
Dim N%
    N = Sz(CnoAy)
If Sz(OAv) <> N Then Er "{CnoAy-Size} is diff from {OAp-Size}", N, Sz(OAv)
Dim J%
For J = 0 To N - 1
    Dim V
        V = Dt.DrAy(R, CnoAy(J))
    OAp(J) = V   '<====
Next
End Sub

Function Dt_CnoAy(Dt As TDt, FldNmAy$()) As Integer()
Dim I&()
    I = Ay_IdxAy(Dt.FldNmAy, FldNmAy)
Dim U%
    U = UBound(I)
Dim O%()
    ReDim O(U)
Dim J%
For J = 0 To U
    O(J) = I(J)
    If I(J) = -1 Then Er "{FldNm} in given {FldNmAy} not in {Table}-{Flds}", FldNmAy(J), Join(Ay_Quote(FldNmAy, "[]")), Dt.Nm, Join(Dt.FldNmAy, " ")
Next
Dt_CnoAy = O
End Function

Function Dt_CnoAy_ByFldNmAy(Dt As TDt, FldNmAy$()) As Integer()
Dt_CnoAy_ByFldNmAy = Dt_CnoAy(Dt, FldNmAy)
End Function

Function Dt_CnoAy_ByFldNmLvs(Dt As TDt, FldNmLvs$) As Integer()
Dt_CnoAy_ByFldNmLvs = Dt_CnoAy(Dt, SplitLvs(FldNmLvs))
End Function
