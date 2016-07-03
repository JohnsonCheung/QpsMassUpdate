Attribute VB_Name = "CfgCpyFmCfgFxToPgmFx"
Option Explicit
' Read 3-CfgFx-Ws and Write 2-PgmFx-Ws
' CfgFx:
'   #1 Ws-CstChr = ChrCd | CtlTyStr MustInp
'   #2 Ws-Ctl    = ChrCd | ChrNm CstGp CstEle
'   #3 Ws-ChrV   = ChrCd ChrValCd | ChrValNm DropDownSel
'   Note: All Ws-CstChr will be selected.  It holds some ChrCd used in cost-analysis
'         Some Ws-Ctl-record-of-ChrCd found in Ws-CstChr will be selected.  It holds all Chr used in QPS
'         Some Ws-ChrV-record-of-ChrCd found in Ws-CstChr will be selected.
' PgmFx:
'   Ws-Ctl = CharCode | CostGp CostEle CharName IsMulti IsMust CtlType
'   Ws-ChrV = CharName | CharValueName CharValueCode
'   Ws-VdtChrCd = ChrCd
' PgmFx-Ws{Ctl,ChrV,VdtChrCd} will be used in a class-ChrDefInf

'===== (Oup=(Ctl ChrV CstChr)==================
Private Type TMsg
    Sht As String
    Msg As String
End Type
Private Type TOup
    Dt As TDt
    Ws As Worksheet
End Type
Const ZCfgFx_WsCstChr_Nm = "Cost Chr"
Const ZCfgFx_WsCtl_Nm = "Ctl"
Const ZCfgFx_WsChrV_Nm = "ChrV"

Private Sub AA()

End Sub

Private Sub Do_CfgCpy__Tst()
Do_CfgCpy
End Sub
Sub Do_CfgCpy()
'#Description
'From CfgFx-Ws-{CstChr Ctl ChrV} create PgmFx-Ws-{Ctl ChrV}
'#Inp-Tables
'CfgFx-Ws-CstChr = ZCfgFx_WsCstChr_FldAy = (ChrCd CstGp CstEle | ChrNm)  ' It has 1 record more than Ctl due to the PK = (ChrCd CstGp CstEle)
'CfgFx-Ws-Ctl    = ZCfgFx_WsCtl   _FldAy = (ChrCd | CtlTyStr MustInp)
'CfgFx-Ws-ChrV   = ZCfgFx_WsChrV  _FldAy = (ChrCd ChrValNm | ChrValCd DropDownSel)
'##Table-Notes
'CfgFx-Ws-CstChr is a subset of CfgFx-Ws-Ctl by means of ChrCd
'                   record of ChrCd means they are selected to be used in Cost-Analysis (So the name is is CstChr - Chr using in Cost-Analysis)
'                   It has records of CtlTyStr = ( .. | .. | .. )
'                   Not all them is required to be copied to PgmFx-Ctl.
'                   Only those CtlTyStr-{} is required to be copied to PgmFx-Ws-Ctl
'##Fields-Notes
'CfgFx-WsCtl-CtlTyStr = (Choose | Input | C)
'CfgFx-WsCtl-MustInp = (..)
'CfgFx-WsChrV-DropDownSel = (..)
'#Oup-Tables
'PgmFx-WsCtl  = (Chr
'PgmFx-WsChrV = (ChrCd ChrValNm | ChrValCd) ' It is subset of CfgFx-Ws-ChrV.  For ChrCd in PgmFx-WsCtl-ChrCd
'##Tables-Notes
'PgmFx-WsCtl contains only those ChrCd has selection of ChrVal.  That means CtlTyStr-{Choose MulitpleValue}
'                          those ChrCd has no selection, example, CtlTyStr-{Input XX} will not be included.
'##Fields-Notes
'
ZCfgFx_DoAssert_3Ws

Dim O_W1 As Worksheet
Dim O_W2 As Worksheet
Dim O_W3 As Worksheet
    Set O_W1 = WsCtl ' @ PgmFx
    Set O_W2 = WsChrV ' @ PgmFx
    Set O_W3 = WsVdtChrCd

Dim O_Sht$
Dim O_Msg$
    ZO_Msg O_Sht, O_Msg
    
Dim O_Dt1 As TDt     ' Ws-Ctl @ PgmFx
Dim O_Dt2 As TDt     ' Ws-ChrV @ PgmFx
Dim O_Dt3 As TDt     ' Ws-VdtChrCd @ PgmFx
    O_Dt1 = ZO_DtCtl

    Dim SelChrCdAy
        SelChrCdAy = Dt_ColAy(O_Dt1, "ChrCd")
    O_Dt2 = ZO_DtChrV(SelChrCdAy)
    O_Dt3 = ZO_DtVdtChrCd
    
Dim O_Cell1 As Range
Dim O_Cell2 As Range
Dim O_Cell3 As Range
    Set O_Cell1 = yWsMassUpd.Range("C5")
    Set O_Cell2 = O_W1.Range("A1")
    Set O_Cell3 = O_W2.Range("A1")

Application.ScreenUpdating = False
O_W1.Cells.Clear     '<== Clear
O_W2.Cells.Clear     '<== Clear
O_W3.Cells.Clear     '<== Clear
Dt_PutWs O_Dt1, O_W1 '<== Put
Dt_PutWs O_Dt2, O_W2 '<== Put
Dt_PutWs O_Dt3, O_W3 '<== Put
Dt_DtaRge(O_Dt1, O_W1).NumberFormat = "@" '<== Fmt to Text
Dt_DtaRge(O_Dt2, O_W2).NumberFormat = "@" '<== Fmt to Text
Dt_DtaRge(O_Dt3, O_W3).NumberFormat = "@" '<== Fmt to Text
O_Cell1.Value = O_Sht   '<= Sht1 @ O_Cell1
Cell_AddCmt O_Cell1, O_Msg, 300, 120    '<== Put Msg @ PgmFx-Ws C5
Cell_AddCmt O_Cell2, O_Msg, 300, 120    '<== Put Msg @ W1.A1
Cell_AddCmt O_Cell3, O_Msg, 300, 120    '<== Put Msg @ W2.A1
O_W1.Visible = xlSheetHidden
O_W2.Visible = xlSheetHidden
O_W3.Visible = xlSheetHidden
Application.ScreenUpdating = True
End Sub

Private Function IsInclude_Ctl(DropDownSel) As Boolean
IsInclude_Ctl = True
If UCase(Trim(DropDownSel)) = "X" Then Exit Function
Dim A%
    A = Val(DropDownSel)
If 1 <= A And A <= 15 Then Exit Function
IsInclude_Ctl = False
End Function

Private Sub ZCfgFx_DoAssert_3Ws()
Dim Er$()
Dim Wb As Workbook
    Set Wb = ZCfgFx_Wb
Wb_Chk_IsWs Wb, ZCfgFx_WsCtl_Nm, Er
Wb_Chk_IsWs Wb, ZCfgFx_WsChrV_Nm, Er
Wb_Chk_IsWs Wb, ZCfgFx_WsCstChr_Nm, Er
Ws_Chk_FldNmAy ZCfgFx_WsCtl, ZCfgFx_WsCtl_FldAy, Er
Ws_Chk_FldNmAy ZCfgFx_WsChrV, ZCfgFx_WsChrV_FldAy, Er
Ws_Chk_FldNmAy ZCfgFx_WsCstChr, ZCfgFx_WsCstChr_FldAy, Er
Assert_AyIsEmpty Er
End Sub

Private Function ZCfgFx_Wb() As Workbook
Dim Wb As Workbook
For Each Wb In Application.Workbooks
    If Wb.Name Like "*(Cfg)*.xlsm" Then Set ZCfgFx_Wb = Wb: Exit Function
Next
Set ZCfgFx_Wb = Workbooks.Open(Fs_CfgFx)
End Function

Private Property Get ZCfgFx_WsChrV() As Worksheet
Set ZCfgFx_WsChrV = Wb_WsOpt(ZCfgFx_Wb, ZCfgFx_WsChrV_Nm)
End Property

Private Property Get ZCfgFx_WsChrV_FldAy() As String()
Const A_W3F_ChrCd = "Classification Codes"
Const A_W3F_ChrValNm = "Sap Char Value Name"
Const A_W3F_ChrValCd = "Sap Char Value Code"
Const A_W3F_DropDownSel = "x=Open to choose" & vbLf & "1..15=DropDown"
ZCfgFx_WsChrV_FldAy = StrAy(A_W3F_ChrCd, A_W3F_ChrValCd, A_W3F_ChrValNm, A_W3F_DropDownSel)
End Property

Private Property Get ZCfgFx_WsCstChr() As Worksheet
Set ZCfgFx_WsCstChr = Wb_WsOpt(ZCfgFx_Wb, ZCfgFx_WsCstChr_Nm)
End Property

Private Property Get ZCfgFx_WsCstChr_FldAy() As String()
Const A_W1F_ChrCd = "Char Code"
Const A_W1F_CstEle = "Element"
Const A_W1F_CstGp = "Group"
Const A_W1F_ChrNm = "Characteristics"
ZCfgFx_WsCstChr_FldAy = StrAy(A_W1F_ChrCd, A_W1F_ChrNm, A_W1F_CstEle, A_W1F_CstGp)
End Property

Private Property Get ZCfgFx_WsCtl() As Worksheet
Set ZCfgFx_WsCtl = Wb_WsOpt(ZCfgFx_Wb, ZCfgFx_WsCtl_Nm)
End Property

Private Property Get ZCfgFx_WsCtl_FldAy() As String()
Const A_W2F_ChrCd = "SAP Charactertistic Code" & vbLf & "(Z_NO_INPUT)"
Const A_W2F_CtlTyStr = "MultipleValue" & vbLf & _
                "Choose" & vbLf & _
                "Input" & vbLf & _
                "NoUpload" & vbLf & _
                "xx = for movement"
Const A_W2F_MustInp = "Must Input?" & vbLf & _
                "NonBlank = Must Input" & vbLf & _
                "Blank = Allow no entry"
ZCfgFx_WsCtl_FldAy = StrAy(A_W2F_ChrCd, A_W2F_CtlTyStr, A_W2F_MustInp)
End Property

Private Function ZO_DtChrV(SelChrCdAy) As TDt
Dim O As TDt
O = Ws_Dt_Sel(ZCfgFx_WsChrV, ZCfgFx_WsChrV_FldAy, "ChrCd ChrValCd ChrValNm DropDownSel", ":ChrCd", eOp_In, SelChrCdAy)
ZO_DtChrV = Dt_Sel(O, SplitLvs("ChrCd ChrValCd ChrValNm"), , "DropDownSel", eOp_Fn, "IsInclude_Ctl")
End Function

Private Function ZO_DtCtl() As TDt
Dim Ws1 As Worksheet
Dim Ws2 As Worksheet
Dim FldAy1$()
Dim FldAy2$()
Dim Dt1 As TDt ' CfgFx-CstChr
Dim Dt2 As TDt ' CfgFx-Ctl
Dim Sz1%
Dim Sz2%
Dim O_DrAy() ' O_DrAy will have same rec as Dt1
              ' Dt1 = ChrCd ChrNm CstEle CstGp      ' It may have 1 rec more than Dt2 due to the PK = ChrCd CstEle CstGp (rec#=39)
              ' Dt2 = ChrCd CtlTyStr MustInp        ' (rec#=38)
              ' O_DrAy = ChrCd ChrNm CstEle CstGp | IsMulti IsMust CtlTyStr
              '          Part1 is from Dt1.
              '          Part2 is from Dt2.
              '          Join Key is ChrCd
              ' Note:
              ' Dt2.MustInp has value inlist (Y or Blank)  (See: Ws-Ctl-of-CtlFx, column-{:MustInp})
              ' O_DrAy.IsMulti = Dt2.CtlTyStr = "MultipleValue"  (See: Ws-Ctl-of-CtlFx, column-{:CtlTyStr}-{A_W2CtlTyStr}
              ' O_DrAy.IsMust  = Dt2.MustInp = "Y"
Dim NR&
Dim Dic As New Dictionary  ' Dic_OfChrCToDr2Idx
Dim J&
Dim DrAy2()
Dim DrAy1()
Dim I&
Dim Dr()
Dim O As TDt
Dim Dr1()
Dim ChrCd$
Dim Dt2Idx&
Dim Dr2()
Dim Dr2MustInp$
Dim Dr2CtlTyStr$
Dim Part2_CtlTyStr$
Dim Part2_IsMulti As Boolean
Dim Part2_IsMust As Boolean

Set Ws1 = ZCfgFx_WsCstChr
Set Ws2 = ZCfgFx_WsCtl
FldAy1 = ZCfgFx_WsCstChr_FldAy
FldAy2 = ZCfgFx_WsCtl_FldAy

Dt1 = Ws_Dt_Sel(Ws1, FldAy1, "ChrCd ChrNm CstEle CstGp")
Dt2 = Ws_Dt_Sel(Ws2, FldAy2, "ChrCd CtlTyStr MustInp", ":ChrCd", eOp_In, Dt_ColAy(Dt1, "ChrCd"))

Sz1 = Sz(Ay_Distinct(Dt_ColAy(Dt1, "ChrCd")))
Sz2 = Sz(Dt2.DrAy)
If Sz1 <> Sz2 Then Er "{Sz1} Count-of-Dist-ChrCd in (B_O_Dt1)-or-(CostEle) should = {Sz2} (B_O_Dt2)-or-(Ctl)", Sz1, Sz2
NR = Sz(Dt1.DrAy)
For J = 0 To UB(Dt2.DrAy)
    Dic.Add Dt2.DrAy(J)(0), J  '<==   (0) is for ChrCd
Next

ReDim O_DrAy(NR - 1)
DrAy1 = Dt1.DrAy
DrAy2 = Dt2.DrAy
For I = 0 To NR - 1
    Dr1 = DrAy1(I)
    ChrCd = Dr1(0)
    Dt2Idx = Dic(ChrCd)
    Dr2 = DrAy2(Dt2Idx)
    Dr2MustInp = Dr2(2)
    Dr2CtlTyStr = Dr2(1)
    
    Part2_CtlTyStr = UCase(Dr2CtlTyStr)
    Part2_IsMulti = Part2_CtlTyStr = "MULTIPLEVALUE"
    Part2_IsMust = UCase(Dr2MustInp) = "Y"

    Dr = DrAy1(I)
    ReDim Preserve Dr(6)
    Dr(4) = Part2_IsMulti
    Dr(5) = Part2_IsMust
    Dr(6) = Part2_CtlTyStr
    O_DrAy(I) = Dr     '<==
Next
O.Nm = ""
O.DrAy = O_DrAy
O.FldNmAy = SplitLvs("ChrCd ChrNm CstGp CstEle IsMulti IsMust CtlTyStr")
ZO_DtCtl = ZO_DtCtl__KeepOnly_CHOOSE_and_MULTIPLEVALUE_or_Must(O)
End Function

Private Function ZO_DtCtl__KeepOnly_CHOOSE_and_MULTIPLEVALUE_or_Must(Dt As TDt) As TDt
Dim DrAy()
Dim O_DrAy() ' Remove DO_DrAy()->CltStr not in (CHOOSE MULTIVALUE)
Dim I%
Dim O As TDt
Dim Dr()
Dim CtlTyStr$
Dim CtlTyStr_Cno%
Dim IsMust_Cno%
Dim IsMust$
CtlTyStr_Cno = Dt_Cno(Dt, "CtlTyStr")
IsMust_Cno = Dt_Cno(Dt, "IsMust")

DrAy = Dt.DrAy
Erase O_DrAy
For I = 0 To UB(DrAy)
    Dr = DrAy(I)
    CtlTyStr = Dr(CtlTyStr_Cno)
    Select Case UCase(CtlTyStr)
    Case "CHOOSE", "MULTIPLEVALUE": Push O_DrAy, Dr '<===
    Case Else
        If Dr(IsMust_Cno) Then Push O_DrAy, Dr '<===
    End Select
Next
O.Nm = Dt.Nm
O.FldNmAy = Dt.FldNmAy
O.DrAy = O_DrAy
ZO_DtCtl__KeepOnly_CHOOSE_and_MULTIPLEVALUE_or_Must = O
End Function

Private Function ZO_DtVdtChrCd() As TDt
Const A_W1F_ChrCd = "Char Code"
ZO_DtVdtChrCd = Ws_Dt_Sel(ZCfgFx_WsCstChr, StrAy(A_W1F_ChrCd), "ChrCd")
End Function

Private Sub ZO_Msg(O_Sht$, O_Msg$)
Const ToWs1 = "Ctl"
Const ToWs2 = "ChrV"
Const FmWs1 = "Cost Chr"
Const FmWs2 = "Ctl"
Const FmWs3 = "ChrV"
Const FmCol1 = "Col1"
Const FmCol2 = "Col2"
Const FmCol3 = "Col3"
Dim I As FInfo
I = Ffn_FInfo(ZCfgFx_Wb.FullName)
Dim Pth$, Sz&, Tim As Date, Fn$
Fn = Ffn_Fn(I.Ffn)
Pth = Ffn_Pth(I.Ffn)
Sz = I.Sz
Tim = I.Tim

O_Sht = "2-hidden-Ws in this program are updated @ " & Now
O_Msg = Fmt_QQ(O_Sht & vbLf & _
    "From-Cfg-Xls-File-Name(?)" & vbLf & _
    "File-Path(?)" & vbLf & _
    "File-Size(?)" & vbLf & _
    "File-Time(?)" & vbLf & _
    "From-Ws1(?).and.(?).and.(?)" & vbLf & _
    "To-2-Hidden-Ws(?).and.(?)", _
    Fn, Pth, Sz, Tim, FmWs1, FmWs2, FmWs3, ToWs1, ToWs2)
End Sub

