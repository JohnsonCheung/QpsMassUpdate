Attribute VB_Name = "DtaChg__Tst"
Option Explicit

Private Sub DtaChg_DoCrt_DtaChgWs_and_PaintWrkWs__Tst()
DtaChg_DoCrt_Ws_and_PaintWrkWs Src_DtaChg, Src_Wb
End Sub

Private Sub DtaChg_KeyDta__Tst()
Dim A() As KeyDta
A = DtaChg_KeyDta(Src_DtaChg)
Stop
End Sub

Private Sub DtaChg_PaintYellow__Tst()
Dim W As Worksheet
    Set W = Src.Wrk.Ws
Dim Y As YellowAdr
    Y = DtaChg_YellowAdr(Src_DtaChg)
    
FmtColor_DoPaint_Yellow W, Y
End Sub

Private Sub TDtaChg__Tst()
Dim Act() As TDtaChg
Act = TDtaChg(Src)
Stop
End Sub
