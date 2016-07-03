Attribute VB_Name = "SrcInf"
Option Explicit

Property Get Src() As TSrc
Dim Wb As Workbook
    Set Wb = Src_Wb
Dim WrkWs As Worksheet
Dim OrgWs As Worksheet
    Set WrkWs = Wb.Sheets(WrkWsNm)
    Set OrgWs = Wb.Sheets(OrgWsNm)
Dim O As TSrc
    O.Wrk = TWsInf(WrkWs)
    O.Org = TWsInf(OrgWs)
Src = O
End Property

Property Get Src_DtaChg() As TDtaChg()
Src_DtaChg = TDtaChg(Src)
End Property

Function Src_DtaEr() As TDtaErOpt
Src_DtaEr = TDtaEr(Src)
End Function

Property Get Src_Fx$()
Src_Fx = Src_Wb.FullName
End Property

Property Get Src_Wb() As Workbook
'Look up a MassUpdWb and return it by this way:
'- Scan Workbooks,
'  If =1 Wb with [*(Mass Update).xlxs], return it.
'  If >1 Wb with , push error and return nothing
'  scan DdtaPth for [*(Mass Update).xlxs], open and return the first one.
'  push error and return nothing
Dim N%
Dim O As Workbook
    Set O = Wb_Lik(Fs_ImpFxLikStr, N)
    
If N = 1 Then
    Set Src_Wb = O
    Exit Property
End If
If N > 1 Then Er "More than one Wb of name {like} are open.  Keep only one opened.", Fs_ImpFxLikStr
Dim Fx$
    Fx = Fs_ImpFx
Set Src_Wb = Application.Workbooks.Open(Fx)
End Property

Private Sub AssertWb()
Dim Wb As Workbook
    Set Wb = Src_Wb
Wb_AssertWsExist Wb, "Working"
Wb_AssertWsExist Wb, "Original"
Ws_AssertSingleListObj Wb.Sheets("Working")
Ws_AssertSingleListObj Wb.Sheets("Original")
'Select Case Er
'Case eWsEr.eMoreThanOneListObjInOrgWs: O = "More than one List-Object in worksheet[Original]"
'Case eWsEr.eMoreThanOneListObjInWrkWs: O = "More than one List-Object in worksheet[Working]"
'Case eWsEr.eMoreThanOneMassUpdWbOpn: O = "More than one [Mass Update Data] workbook is opened.  Keep only one open"
'Case eWsEr.eNoListObjInOrgWs: O = "In Worksheet[Original], there is no List-Object"
'Case eWsEr.eNoListObjInWrkWs: O = "In Worksheet[Working], there is no List-Object"
'Case eWsEr.eNoMassUpdWb: O = "No [Mass Update Data] workbook is opened or found in folder[" & Wb_Pth(yWbMassUpd) & "]"
'Case eWsEr.eNoOrgWs: O = "No Worksheet[Original] is found in current [Mass Update Data] workbook"
'Case eWsEr.eNoWrkWs: O = "No Worksheet[Working] is found in current [Mass Update Data] workbook"
'Case Else: Stop
'End Select
'WsErMsg = O
End Sub
