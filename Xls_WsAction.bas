Attribute VB_Name = "Xls_WsAction"
Option Explicit

Sub Ws_ClrOleObjs(Ws As Worksheet)
Dim J%
For J% = Ws.OLEObjects.Count To 1 Step -1
    Dim iOleObj As OLEObject: Set iOleObj = Ws.OLEObjects(J)
    iOleObj.Delete
Next
End Sub

Sub Ws_TblBrw(Ws As Worksheet)
Dt_Brw Ws_Dt(Ws)
End Sub

Private Sub Ws_TblBrw__Tst()
Ws_TblBrw ErWsV3
End Sub
