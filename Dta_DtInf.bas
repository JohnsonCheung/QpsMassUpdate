Attribute VB_Name = "Dta_DtInf"
Option Explicit

Type TDt
    Nm As String
    FldNmAy() As String
    DrAy() As Variant
End Type

Function Dt_Cno%(Dt As TDt, FldNm$)
Dim A%
A = Ay_Idx(Dt.FldNmAy, FldNm)
If A = -1 Then Er "{FldNm} not in {Dt} of {FldNmAy}", FldNm, Dt.Nm, Join(Dt.FldNmAy)
Dt_Cno = A
End Function

Function DrAy_Sqv(DrAy(), NC%) As Variant
Dim O()
    ReDim O(1 To Sz(DrAy), 1 To NC)
Dim R&
For R = 1 To Sz(DrAy)
    Dim Dr
        Dr = DrAy(R - 1)
    Dim C%
    For C = 1 To Min(NC, Sz(Dr))
        O(R, C) = Dr(C - 1)
    Next
Next
DrAy_Sqv = O
End Function

Function Dt_ColAy(Dt As TDt, ColNm$) As Variant()
'Return a column of Name {ColNm} in {Dt}
Dim U&
    U = UB(Dt.DrAy)
If U = -1 Then Exit Function
Dim Cno%
    Cno = Ay_Idx(Dt.FldNmAy, ColNm)
    If Cno = -1 Then Er "Given {ColNm} not exist in {FldNmAy} of {Dt}", ColNm, Dt.FldNmAy, Dt.Nm
Dim O()
    ReDim O(U)
    Dim R&
    For R = 0 To UB(Dt.DrAy)
        Dim Dr()
            Dr = Dt.DrAy(R)
        If Cno <= UB(Dr) Then
            O(R) = Dr(Cno)
        End If
    Next
Dt_ColAy = O
End Function

Function Dt_DtaRge(Dt As TDt, Ws As Worksheet) As Range
Dim R2&
    R2 = Sz(Dt.DrAy) + 1
Dim C2%
    C2 = Sz(Dt.FldNmAy)
Set Dt_DtaRge = Ws_RCRC(Ws, 2, 1, R2, C2)
End Function

Sub Dt_PutWs(Dt As TDt, Ws As Worksheet)
Dim C2%
    C2 = Sz(Dt.FldNmAy)
    If C2 = 0 Then Er "Size of Dt.FldNmAy is zero"
Dim H As Range
    Set H = Ws_RCC(Ws, 1, 1, C2)

H.Value = Dt.FldNmAy    '<== Put HD
Cell_PutSqv Ws_RC(Ws, 2, 1), Dt_Sqv(Dt) '<== Put DtaRge

Dim O_AllRge As Range
    Set O_AllRge = Ws_RCRC(Ws, 1, 1, Sz(Dt.DrAy) + 1, C2)

O_AllRge.Columns.AutoFit '<== AutoFit
Ws.ListObjects.Add(xlSrcRange, O_AllRge, , xlYes).Name = "Dt_" & Dt.Nm  '<== Create ListObject
End Sub

Function Dt_Sqv(Dt As TDt) As Variant
Dt_Sqv = DrAy_Sqv(Dt.DrAy, Sz(Dt.FldNmAy))
End Function

Function TDt(Nm$, FldNmAy$(), DrAy()) As TDt
Dim O As TDt
O.Nm = Nm
O.FldNmAy = FldNmAy
O.DrAy = DrAy
TDt = O
End Function

Private Sub Dt_PutWs__Tst()
Dim Dt As TDt
    Dt = Ws_Dt(ErWsV3)
Dt_PutWs Dt, yWbMassUpd.Sheets("Sheet2")

End Sub
