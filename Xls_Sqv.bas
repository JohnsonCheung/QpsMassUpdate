Attribute VB_Name = "Xls_Sqv"
Option Explicit

Sub Sqv_Brw(HdSqv, SrcSqv)
Dim Ws As Worksheet
Set Ws = Ws_New
Cell_PutSqv Ws.Range("A1"), HdSqv
Cell_PutSqv Ws.Range("A2"), SrcSqv
Ws_Wb(Ws).Activate
Ws.Activate
End Sub

Function Sqv_DrAy(Sqv) As Variant()
Dim NRow&
    NRow = UBound(Sqv, 1)
If NRow = 0 Then Exit Function
Dim ODrAy()
    ReDim ODrAy(NRow - 1)
Dim R&
Dim Sqv_NC%
    Sqv_NC = UBound(Sqv, 2)
For R = 1 To NRow
    Dim NC%
        For NC = Sqv_NC To 1 Step -1
            If Not IsEmpty(Sqv(R, NC)) Then Exit For
        Next
    Dim Dr()
        Erase Dr
        If NC > 0 Then
            ReDim Dr(NC - 1)
            Dim C%
            For C = 0 To NC - 1
                Dr(C) = Sqv(R, C + 1)
            Next
        End If
    ODrAy(R - 1) = Dr
Next
Sqv_DrAy = ODrAy
End Function

Function Sqv_GetDr_Base1(Sqv, R&) As Variant()
Dim U&, J%
U = UBound(Sqv, 2)
ReDim O(1 To U)
For J = 1 To U
    O(J) = Sqv(R, J)
Next
Sqv_GetDr_Base1 = O
End Function

Sub Sqv_PutDr_Base1(Sqv, R&, Dr())
Dim J%
For J = 1 To UBound(Dr)
    Sqv(R, J) = Dr(J)
Next
End Sub

Function Sqv_Row1_ToAy(Sqv, OAy)
Dim U&
    U = UBound(Sqv, 2)
ReDim OAy(U - 1)
Dim J%
For J = 1 To U
    OAy(J - 1) = Sqv(1, J)
Next
Sqv_Row1_ToAy = OAy
End Function

Function Sqv_TrimStr(Sqv)
Dim R&, C&
For R = 1 To UBound(Sqv, 1)
    For C = 1 To UBound(Sqv, 2)
        If VarType(Sqv(R, C)) = vbString Then Sqv(R, C) = Trim(Sqv(R, C))
    Next
Next
Sqv_TrimStr = Sqv
End Function
