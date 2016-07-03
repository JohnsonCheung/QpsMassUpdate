Attribute VB_Name = "DoGenCmpRpt"
Option Explicit
Private A_NewQDte As Date
Private ZMassUpd_Fx$
Private ZNewQQry_Fx$
Private ZCmpRpt_Fx$
Private ZMassUpd_Wb As Workbook
Private ZNewQQry_Wb As Workbook
Private ZMassUpd_Ws As Worksheet
Private ZNewQQry_Ws As Worksheet
Private ZCmpRpt_Wb As Workbook
Const ZNewQQry_NewWsNm = "AA"
Const ZNewQQry_WsNm = "AA"
Const ZMassUpd_NewWsNm = "AA"
Const ZMassUpd_WsNm = "AA"

Sub Do_GenCmpRpt(NewQDte As Date)
A_NewQDte = NewQDte
ZMassUpd_Fx = ZZMassUpd_Fx
ZNewQQry_Fx = ZZNewQQry_Fx
ZCmpRpt_Fx = ZZCmpRpt_Fx
ZMassUpd_DoAssert_Fx
ZNewQQry_DoAssert_Fx
ZMassUpd_DoOpn_Fx   ' Wb & Ws are set
ZNewQQry_DoOpn_Fx   ' Wb & Ws are set
ZCmpRpt_DoNew_Wb    ' Wb is set
ZCmpRpt_DoCpy_Fm_MassUpd
ZCmpRpt_DoCpy_Fm_NewQQry
ZNewQQry_DoClose
ZMassUpd_DoClose
ZCmpRpt_DoGen_Cmp
ZCmpRpt_DoSavAndClose
Application.Quit
End Sub

Private Sub ZCmpRpt_DoCpy_FmWs(FmWs As Worksheet, NewWsNm$)
Dim ToWb As Workbook
    Set ToWb = ZCmpRpt_Wb

Dim ToWs As Worksheet
    Set ToWs = ToWb.Sheets(ToWb.Sheets.Count)

FmWs.Copy , ToWs

Dim NewWs As Worksheet
    Set NewWs = ToWb.Sheets(ToWb.Sheets.Count)

NewWs.Name = NewWsNm

End Sub

Private Sub ZCmpRpt_DoCpy_Fm_MassUpd()
ZCmpRpt_DoCpy_FmWs ZMassUpd_Ws, "AA"
End Sub

Private Sub ZCmpRpt_DoCpy_Fm_NewQQry()
ZCmpRpt_DoCpy_FmWs ZNewQQry_Ws, ZNewQQry_NewWsNm
End Sub

Private Sub ZCmpRpt_DoGen_Cmp()

End Sub

Private Sub ZCmpRpt_DoNew_Wb()
Set ZCmpRpt_Wb = Application.Workbooks.Add
End Sub

Private Sub ZCmpRpt_DoSavAndClose()
ZCmpRpt_Wb.SaveAs ZCmpRpt_Fx
ZCmpRpt_Wb.Close
End Sub

Private Sub ZMassUpd_DoAssert_Fx()

End Sub

Private Sub ZMassUpd_DoClose()
ZMassUpd_Wb.Close False
End Sub

Private Sub ZMassUpd_DoOpn_Fx()
Set ZMassUpd_Wb = Application.Workbooks.Open(ZMassUpd_Fx)
Set ZMassUpd_Ws = ZMassUpd_Wb.Sheets(ZMassUpd_WsNm)
End Sub

Private Sub ZNewQQry_DoAssert_Fx()

End Sub

Private Sub ZNewQQry_DoClose()
ZNewQQry_Wb.Close False
End Sub

Private Sub ZNewQQry_DoOpn_Fx()
Set ZNewQQry_Wb = Application.Workbooks.Open(ZNewQQry_Fx)
Set ZNewQQry_Ws = ZNewQQry_Wb.Sheets(ZNewQQry_WsNm)
End Sub

Private Property Get ZZCmpRpt_Fx$()

End Property

Private Property Get ZZMassUpd_Fx$()

End Property

Private Property Get ZZNewQQry_Fx$()

End Property
