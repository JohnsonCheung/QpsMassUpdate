Attribute VB_Name = "MacroErWsSelChgV4SelInSep"
Option Explicit
Private Enum eWbTy
    eOther = 0
    ePgm = 1
    eCorr = 2
    eSel = 3
End Enum

Private Type VBar
    R1 As Long
    R2 As Long
    C As Integer
End Type
Private A_Cell As Range
Private A_Corr_LastCell As Range
Private A_Corr_CurCell As Range
Private A_Sel_StopSelChg As Boolean
Private Const ZSelWb_Nm$ = "QPS-MassUpd-Selection.xlsx"

Sub ErWs_WsSelChg_V4SelInSep(Cell As Range)
If Application.CutCopyMode <> 0 Then Exit Sub
If ZSel_StopSelChg Then Exit Sub
Set A_Cell = Cell
ZSel_StopSelChg = True
Application.ScreenUpdating = False
A_Main
ZSel_StopSelChg = False
Application.ScreenUpdating = True
End Sub

Private Sub A_Main()
'MsgBox "CellAdr=" & A_Cell.Address & "; IsMulit=" & ZIsMulti & "; Adr=" & ZMulti_Cell.Address
ZCurCell_Assert
ZCorr_DoSet_CurCurrCell
If ZCorr_IsNoEr Then Exit Sub
ZSelWb_DoEnsure
'If ZIsCurCell_InWsSel Then MsgBox "In Selection Ws"
'If ZIsCurCell_InWsCorr Then MsgBox "In Corr Ws"
'Exit Sub
If ZIsCurCell_InSide_CorrVBar Then
    Set A_Corr_CurCell = A_Cell
    ZCorr_DoClr_LastCellBdr
    ZSel_DoBld
    Exit Sub
End If
If ZIsCurCell_InSide_SelValVBar Then
    If ZIsMulti Then
        ZSelCell_DoToggle
    Else
        ZSelCell_DoSelect
    End If
    ZCorr_DoSet_CellVal
    ZTar_DoSet_CellVal
Else
    ZCorr_DoClr_LastCellBdr
    ZSel_DoClr
End If
End Sub

Private Sub Class_Initialize()
ZCorr_DoSet_ColAsTxt
ZSel_DoSet_ColAsTxt
End Sub

Private Sub ZCell_DoHighlight(Cell As Range)
Rge_RC(Cell, 1, 1).Interior.Color = rgbYellow
End Sub

Private Sub ZCell_DoSetVal(Cell As Range, ValStr$)
With Cell
    With .Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    .Font.ColorIndex = xlAutomatic
    .Font.TintAndShade = 0
    .Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone
    .NumberFormat = "@"
    .Value = ValStr
End With
End Sub

Private Sub ZCell_DoUnlight(Cell As Range)
Rge_RC(Cell, 1, 1).Interior.Pattern = xlNone
End Sub

Private Property Get ZCharName_CharValNameAy() As String()
ZCharName_CharValNameAy = ChrDefInf.ChrNm_ChrValNmAy(ZCharName_CurVal)
End Property

Private Property Get ZCharName_Cno%()
ZCharName_Cno = ZZFndCno("CharName")
End Property

Private Property Get ZCharName_CurVal$()
ZCharName_CurVal = Ws_RC(ZCurCell_Ws, ZR, ZCharName_Cno).Value
End Property

Private Function ZCharValName_Cell1() As Range
Dim C%
    C = 1
Dim A As Range
    Set A = Ws_RC(ZSel_Ws, 1, 1)
If IsEmpty(A.Value) Then Exit Function
Set ZCharValName_Cell1 = A
End Function

Private Function ZCharValName_Cell2() As Range
Dim A As Range
    Set A = ZCharValName_Cell1
If IsNothing(A) Then Exit Function
If IsEmpty(A.Range("A2").Value) Then
    ZCharValName_Cell2 = A
    Exit Function
End If
Set ZCharValName_Cell2 = A.End(xlDown)
End Function

Private Sub ZCharValName_DoHighlight_FmCorrVal()
Dim C As Range
    Set C = ZCorr_CurCell
If IsNothing(C) Then Exit Sub
Dim A$
    A = ZCorr_CurCell.Value
Dim B$()
    B = Split(A, vbLf)
Dim R As Range
    Set R = ZCharValName_Rge
If R.Count = 0 Then Exit Sub
Dim Cell As Range
For Each Cell In R
    Dim CellVal$
        CellVal = Cell.Value
    If Ay_Has(B, CellVal) Then ZCell_DoHighlight Cell '<====
Next
End Sub

Private Sub ZCharValName_DoSel_FirstHighlightCell()
ZSel_Wb.Activate
ZCharValName_FirstHighlightCell.Select
End Sub

Private Sub ZCharValName_DoUnlight()
Dim Ay() As Range, J%
ZCharValName_Rge.Interior.Pattern = xlNone
End Sub

Private Property Get ZCharValName_FirstHighlightCell() As Range
Dim Rge As Range
    Set Rge = ZCharValName_Rge
Dim R As Range
For Each R In Rge
    If ZIsCell_Highlight(R) Then Set ZCharValName_FirstHighlightCell = R: Exit Property
Next
Set ZCharValName_FirstHighlightCell = ZSel_Ws.Range("A1")
End Property

Private Property Get ZCharValName_Rge() As Range
Dim B As VBar
B = ZSel_ValVBar
Set ZCharValName_Rge = Ws_CRR(ZSel_Ws, B.C, B.R1, B.R2)
End Property

Private Function ZColA_R2&()
ZColA_R2 = Ws_RC(ZCurCell_Ws, 1, 1).End(xlDown).Row
End Function

Private Function ZCorr_Cno%()
ZCorr_Cno = ZZFndCno("Correction")
End Function

Private Property Get ZCorr_CurCell() As Range
Set ZCorr_CurCell = A_Corr_CurCell
End Property

Private Sub ZCorr_DoClr_LastCellBdr()
If IsNothing(A_Corr_LastCell) Then Exit Sub
Cell_ClrBdr A_Corr_LastCell
End Sub

Private Sub ZCorr_DoSavAs_LastCell()
Set A_Corr_LastCell = A_Cell
End Sub

Private Sub ZCorr_DoSet_CellVal()
Dim A$
    A = ZSelected_NewValStr
Dim Cell As Range
    Set Cell = ZCorr_LastCell
ZCell_DoSetVal Cell, A
Cell.EntireColumn.AutoFit
End Sub

Private Sub ZCorr_DoSet_ColAsTxt()
Ws_C(ZCurCell_Ws, ZCorr_Cno).NumberFormat = "@"
End Sub

Private Sub ZCorr_DoSet_CurCurrCell()
If ZIsCurCell_InSide_CorrVBar Then Set A_Corr_CurCell = A_Cell
End Sub

Private Function ZCorr_IsCorrWb(Wb As Workbook) As Boolean
ZCorr_IsCorrWb = Wb_IsWs(Wb, "Error")
End Function

Private Property Get ZCorr_IsNoEr() As Boolean
Dim Ws As Worksheet
    Set Ws = ZCorr_Ws
Dim R2&
    R2 = Ws.Range("A1").End(xlDown).Row
Dim Rge As Range
    Set Rge = Ws_CRR(Ws, 1, 2, R2)
Dim R As Range
For Each R In Rge
    If R.Value = "Empty Char" Then Exit Property
    If R.Value = "Invalid Char Val" Then Exit Property
Next
ZCorr_IsNoEr = True
End Property

Private Function ZCorr_LastCell() As Range
Set ZCorr_LastCell = A_Corr_LastCell
End Function

Private Property Get ZCorr_LastCellAdr$()
ZCorr_LastCellAdr = ZCorr_LastCell.Address
End Property

Private Function ZCorr_VBar() As VBar
Dim O As VBar
O.C = ZCorr_Cno
O.R1 = 2
O.R2 = ZColA_R2
ZCorr_VBar = O
End Function

Private Property Get ZCorr_Wb() As Workbook
Dim Wb As Workbook
For Each Wb In Application.Workbooks
    If ZCorr_IsCorrWb(Wb) Then Set ZCorr_Wb = Wb: Exit Property
Next
Er "No workbook with sheet name is [Error]"
End Property

Private Function ZCorr_Ws() As Worksheet
Set ZCorr_Ws = ZCorr_Wb.Sheets("Error")
End Function

Private Sub ZCurCell_Assert()
Dim Nm$
    Nm = A_Cell.Worksheet.Name
If Nm = "Selection" Then Exit Sub
If Nm = "Error" Then Exit Sub
Er "The {WsNm} of {Cell} should be [Selection] or [Error]", Nm, A_Cell.Address
End Sub

Private Sub ZCurCell_DoBdr()
A_Cell.BorderAround XlLineStyle.xlContinuous
End Sub

Private Sub ZCurCell_DoScrollToTop()
ActiveWindow.ScrollRow = A_Cell.Row
End Sub

Private Property Get ZCurCell_PossibleSelectionSqv()
Dim A$()
    A = ZCharName_CharValNameAy
Dim N&
    N = Sz(A)
If N = 0 Then Exit Property
ReDim O(1 To N, 1 To 1)
    Dim J%
    For J = 1 To N
        O(J, 1) = A(J - 1)
    Next
ZCurCell_PossibleSelectionSqv = O
End Property

Private Function ZCurCell_Wb() As Workbook
Set ZCurCell_Wb = ZCurCell_Ws.Parent
End Function

Private Function ZCurCell_Ws() As Worksheet
Set ZCurCell_Ws = A_Cell.Worksheet
End Function

Private Function ZIsCell_Highlight(Cell As Range) As Boolean
ZIsCell_Highlight = Cell.Interior.Color = rgbYellow
End Function

Private Property Get ZIsCurCell_InSide_CorrVBar() As Boolean
If Not ZIsCurCell_InWsCorr Then Exit Function
ZIsCurCell_InSide_CorrVBar = ZIsCurCell_InSide_VBar(ZCorr_VBar)
End Property

Private Property Get ZIsCurCell_InSide_SelValVBar() As Boolean
If Not ZIsCurCell_InWsSel Then Exit Function
ZIsCurCell_InSide_SelValVBar = ZIsCurCell_InSide_VBar(ZSel_ValVBar)
End Property

Private Property Get ZIsCurCell_InSide_SelVbar() As Boolean
If Not ZIsCurCell_InWsSel Then Exit Function
ZIsCurCell_InSide_SelVbar = ZIsCurCell_InSide_VBar(ZSel_VBar)
End Property

Private Property Get ZIsCurCell_InSide_VBar(B As VBar) As Boolean
ZIsCurCell_InSide_VBar = Not ZIsCurCell_OutSide_VBar(B)
End Property

Private Property Get ZIsCurCell_InWsCorr() As Boolean
ZIsCurCell_InWsCorr = Not ZIsCurCell_InWsSel
End Property

Private Property Get ZIsCurCell_InWsSel() As Boolean
If ZSelWb_IsNoSelWb Then Exit Property
ZIsCurCell_InWsSel = ObjPtr(ZSel_Ws) = ObjPtr(A_Cell.Worksheet)
End Property

Private Function ZIsCurCell_OutSide_VBar(B As VBar) As Boolean
Dim R&, C%
    R = A_Cell.Row
    C = A_Cell.Column
ZIsCurCell_OutSide_VBar = True
If B.R1 > R Or R > B.R2 Then Exit Function
If C <> B.C Then Exit Function
ZIsCurCell_OutSide_VBar = False
End Function

Private Function ZIsMulti() As Boolean
Dim O As Boolean
    O = ZMulti_Cell.Value = "Multi"
ZIsMulti = O
End Function

Private Function ZIsMust() As Boolean
ZIsMust = ZMust_Cell.Value = "Must"
End Function

Private Function ZIsSingle() As Boolean
ZIsSingle = Not ZIsMulti
End Function

Private Function ZMulti_Cell() As Range
Set ZMulti_Cell = ZCorr_CurCell.EntireRow.Cells(1, ZMulti_Cno)
End Function

Private Property Get ZMulti_Cno%()
ZMulti_Cno = ZZFndCno("Multi")
End Property

Private Function ZMust_Cell() As Range
Set ZMust_Cell = ZCorr_CurCell.EntireRow.Cells(1, ZMust_Cno)
End Function

Private Property Get ZMust_Cno%()
ZMust_Cno = ZZFndCno("Must")
End Property

Private Property Get ZR&()
ZR = A_Cell.Row
End Property

Private Sub ZSelCell_DoHighlight()
ZCell_DoHighlight A_Cell
End Sub

Private Sub ZSelCell_DoSelect()
ZCharValName_DoUnlight
ZSelCell_DoHighlight
End Sub

Private Sub ZSelCell_DoToggle()
If ZSelCell_IsHighlight Then
    ZSelCell_DoUnlight
Else
    ZSelCell_DoHighlight
End If
End Sub

Private Sub ZSelCell_DoUnlight()
ZCell_DoUnlight A_Cell
End Sub

Private Function ZSelCell_IsHighlight() As Boolean
ZSelCell_IsHighlight = ZIsCell_Highlight(A_Cell)
End Function

Private Sub ZSelWb_DoAddMacro()
Ws_Crt_EvtMth_CallingFn ZSel_Ws, "SelectionChange", "jjMassUpd", , "ErWs_WsSelChg_V4SelInSep"
End Sub

Private Sub ZSelWb_DoEnsure()
Dim A$()
    A = ZSelWb_WbToClose
If Sz(A) > 0 Then
    MsgBox Join(A, vbCrLf), vbInformation, "Please close this workbooks"
    Dim Wb As Workbook
    For Each Wb In Application.Workbooks
        If Wb.Name = A(0) Then Wb.Activate: Exit Sub
    Next
    Exit Sub
End If
If ZSelWb_IsNoSelWb Then
    ZSelWb_DoNew
    ZSelWb_DoAddMacro
    ZWin_DoFmtSel
End If
End Sub

Private Sub ZSelWb_DoNew()
Dim Wb As Workbook
Set Wb = Application.Workbooks.Add
Wb_DltWs Wb, "Sheet3"
Wb_DltWs Wb, "Sheet2"
Wb_Ws(Wb, "Sheet1").Name = "Selection"
Dim Fx$
    Fx = Fso.GetSpecialFolder(2) & "\" & ZSelWb_Nm
Ffn_DltIfExist Fx
Wb.SaveAs Fx
Wb.Close False
Set Wb = Application.Workbooks.Open(Fx)
If Wb.Sheets.Count <> 1 Then Er "{Wb} Sheets {count} should be 1", Wb.Name, Wb.Sheets.Count
If Wb_Ws(Wb, "Selection").CodeName = "" Then Er "CodeName of Selection-Ws of {Wb} is empty", Wb.Name
End Sub

Private Property Get ZSelWb_IsNoSelWb() As Boolean
Dim Wb As Workbook
For Each Wb In Application.Workbooks
    If ZSelWb_WbTy(Wb) = eSel Then Exit Property
Next
ZSelWb_IsNoSelWb = True
End Property

Private Function ZSelWb_IsWbToClose(Wb As Workbook) As Boolean
Select Case ZSelWb_WbTy(Wb)
Case ePgm, eSel, eCorr: Exit Function
End Select
ZSelWb_IsWbToClose = True
End Function

Private Property Get ZSelWb_WbToClose() As String()
Dim O$()
Dim Wb As Workbook
For Each Wb In Application.Workbooks
    If ZSelWb_IsWbToClose(Wb) Then Push O, Wb.Name
Next
ZSelWb_WbToClose = O
End Property

Private Function ZSelWb_WbTy(Wb As Workbook) As eWbTy
Dim O As eWbTy
If Wb_IsWs(Wb, "Error") Then
    O = eCorr
ElseIf Wb.Name = ZSelWb_Nm Then
    O = eSel
ElseIf Wb.FullName = Application.VBE.activeproject.File Then
    O = ePgm
Else
    O = eOther
End If
ZSelWb_WbTy = O
End Function

Private Function ZSel_Cno%()
ZSel_Cno = 1
End Function

Private Function ZSel_ColVBar() As VBar
Dim O As VBar
O.C = ZSel_Cno
O.R1 = 2
O.R2 = Ws_MaxRno(ZCurCell_Ws)
ZSel_ColVBar = O
End Function

Private Sub ZSel_DoBld()
ZSel_DoClr
'Dim A As Dictionary
'    'Set A = Cfg_ChrNmDic_ToChrValNm
'    Dim B()
'    'B = A.Items
    
Dim Sqv
    Sqv = ZCurCell_PossibleSelectionSqv
If IsEmpty(Sqv) Then Exit Sub

Dim N%
    N = UBound(Sqv, 1)
ZSel_DoShw_SelectionRowOnly N

Dim R As Range
    Set R = Ws_CRR(ZSel_Ws, 1, 1, N)
    R.NumberFormat = "@"
    R.VerticalAlignment = XlVAlign.xlVAlignCenter
R.Value = Sqv
ZCharValName_DoHighlight_FmCorrVal
ZCharValName_DoSel_FirstHighlightCell
ZCorr_DoSavAs_LastCell
ZCurCell_DoBdr
ZWin_DoEnsure_OneWb_OneWin
ZWin_DoPosition
'ZSel_DoAct_FirstCell
ZSel_Ws.Columns("A:A").EntireColumn.AutoFit
End Sub

Private Sub ZSel_DoClr()
ZSel_Ws.Cells.Clear
ZSel_Ws.Cells.EntireRow.Hidden = True
End Sub

Private Sub ZSel_DoSet_ColAsTxt()
Ws_C(ZCurCell_Ws, ZSel_Cno).NumberFormat = "@"
End Sub

Private Sub ZSel_DoShw_SelectionRowOnly(NRow%)
Dim Ws As Worksheet
    Set Ws = ZSel_Ws
Ws.Cells.EntireRow.Hidden = False
Dim M&
    M = Ws_MaxRno(Ws)
Ws_RR(Ws, NRow + 1, M).EntireRow.Hidden = True
End Sub

Private Property Get ZSel_R2&()
Dim Ws As Worksheet
    Set Ws = ZSel_Ws
Dim A As Range
    Set A = Ws_RC(Ws, 1, 1)
If IsEmpty(A.Value) Then Exit Property
Dim B As Range
    Set B = Ws_RC(Ws, 2, 1)
If IsEmpty(B.Value) Then ZSel_R2 = 1: Exit Property
ZSel_R2 = A.End(xlDown).Row
End Property

Private Property Get ZSel_StopSelChg() As Boolean
ZSel_StopSelChg = A_Sel_StopSelChg
End Property

Private Property Let ZSel_StopSelChg(V As Boolean)
A_Sel_StopSelChg = V
End Property

Private Function ZSel_VBar() As VBar
Dim O As VBar
O.R1 = 1
O.R2 = ZSel_R2
O.C = ZSel_Cno
ZSel_VBar = O
End Function

Private Function ZSel_ValVBar() As VBar
Dim A As Range
    Set A = ZCharValName_Cell1
    If IsNothing(A) Then Exit Function
Dim B As Range
    Set B = ZCharValName_Cell2
    If IsNothing(B) Then Exit Function

Dim R1&, R2&
    R1 = A.Row
    R2 = B.Row
Dim O As VBar
O.R1 = R1
O.R2 = R2
O.C = ZSel_Cno
ZSel_ValVBar = O
End Function

Private Property Get ZSel_Wb() As Workbook
Dim Wb As Workbook
For Each Wb In Application.Workbooks
    If ZSelWb_WbTy(Wb) = eSel Then Set ZSel_Wb = Wb: Exit Property
Next
End Property

Private Function ZSel_Win() As Window
Dim Ws As Worksheet
    Set Ws = ZCurCell_Ws
Dim W As Window
For Each W In ZCurCell_Wb.Windows
    If W.Activate = Ws Then Set ZSel_Win = W: Exit Function
Next
End Function

Private Property Get ZSel_Ws() As Worksheet
'Sel_Ws always exist, before calling this SelectionChange, before, it is required to add
'The handling method to the Sel_Ws
Set ZSel_Ws = ZSel_Wb.Sheets("Selection")
End Property

Private Property Get ZSelected_CharValNameAy() As String()
Dim A As Range
    Set A = ZCharValName_Rge
If A.Count = 0 Then Exit Property
Dim O$()
    Dim R As Range
    For Each R In A
        If ZIsCell_Highlight(R) Then Push O, R.Value
    Next
ZSelected_CharValNameAy = O
End Property

Private Function ZSelected_NewValStr$()
ZSelected_NewValStr = Join(ZSelected_CharValNameAy, vbLf)
End Function

Private Function ZTar_Cell() As Range
Set ZTar_Cell = ZTar_Ws.Range(ZWrkAdr_CellVal)
End Function

Private Function ZTar_CellVal_Cur$()
ZTar_CellVal_Cur = ZTar_Cell.Value
End Function

Private Sub ZTar_Cell_DoResetColor()
ZTar_Cell.Font.ColorIndex = xlAutomatic
End Sub

Private Sub ZTar_Cell_DoSetRed()
ZTar_Cell.Font.Color = rgbRed
End Sub

Private Sub ZTar_DoSet_CellVal()
If Not ZTar_WsExist Then Exit Sub
ZCell_DoSetVal ZTar_Cell, ZSelected_NewValStr
End Sub

Private Function ZTar_Ws() As Worksheet
Set ZTar_Ws = ZCurCell_Ws.Parent.Sheets("Working")
End Function

Private Property Get ZTar_WsExist() As Boolean
ZTar_WsExist = Wb_IsWs(Ws_Wb(ZCurCell_Ws), "Working")
End Property

Private Function ZWin_1Width#()
ZWin_1Width = 700
End Function

Private Function ZWin_2Left#()
ZWin_2Left = ZWin_1Width + 1
End Function

Private Function ZWin_2Width#()
ZWin_2Width = ZWin_ScnWidth - ZWin_1Width
End Function

Private Sub ZWin_DoEnsure_OneWb_OneWin()
Dim Wb As Workbook
Dim I%
For Each Wb In Application.Workbooks
    Do While Wb.Windows.Count > 1
        I = I + 1
        If I > 10 Then Er "Program Err"
        Wb.Windows(1).Close
    Loop
Next
Dim C%
    C = Application.Windows.Count
If C = 2 Or C = 3 Then Exit Sub
Er "There should 2 or 3 windows, but now it has {n} windows", C
End Sub

Private Sub ZWin_DoFmtSel()

With ZWin_Sel
    .DisplayHeadings = False
    .DisplayGridlines = False
    .DisplayHorizontalScrollBar = False
    .DisplayWorkbookTabs = False
    .Zoom = 75
End With
Dim Ws As Worksheet
    Set Ws = ZSel_Ws
Dim C2%
    C2 = Ws_MaxCno(Ws)
Ws_CC(Ws, 2, C2).Hidden = True
End Sub

Private Sub ZWin_DoPosition()
Dim H#
    H = ZWin_Hgt
    
Dim Wb As Workbook
For Each Wb In Application.Workbooks
    Select Case ZSelWb_WbTy(Wb)
    Case ePgm
        Wb.Windows(1).WindowState = xlMinimized
    Case eSel
        With Wb.Windows(1)
            .WindowState = xlNormal
            .EnableResize = False
            .Top = 1
            .Left = ZWin_1Width + 1
            .Width = ZWin_2Width - 4
            .Height = H
            .Caption = "Selection"
            .EnableResize = False
        End With
    Case eCorr
        With Wb.Windows(1)
            .WindowState = xlNormal
            .Top = 1
            .Left = 1
            .Width = ZWin_1Width
            .Height = H
        End With
    Case Else
        Er "Program Err: The {Wb} in openned is not expected [Corr, Sel, Pgm}", Wb.Name
    End Select
Next
End Sub

Private Function ZWin_Height#()
ZWin_Height = ZWin_ScnHeight#
End Function

Private Property Get ZWin_Hgt#()
Dim J%
For J = 2 To Application.Windows.Count
    Application.Windows(J).EnableResize = True
    Application.Windows(J).WindowState = xlMinimized
Next
Dim B As Boolean
    B = Application.Windows(1).EnableResize
    Application.Windows(1).EnableResize = True
Application.Windows.Arrange xlArrangeStyleHorizontal
Application.Windows(1).EnableResize = B
ZWin_Hgt = Application.Windows(1).Height
End Property

Private Function ZWin_ScnHeight#()
ZWin_ScnHeight = Application.Height
End Function

Private Function ZWin_ScnWidth#()
ZWin_ScnWidth = Application.Width
End Function

Private Property Get ZWin_Sel() As Window
Set ZWin_Sel = ZSel_Wb.Windows(1)
End Property

Private Function ZWrkAdr_Cell() As Range
Set ZWrkAdr_Cell = ZCurCell_Ws.Cells(ZCharValName_Cell1.Row, ZWrkAdr_Cno)
End Function

Private Function ZWrkAdr_CellVal$()
ZWrkAdr_CellVal = ZWrkAdr_Cell.Value
End Function

Private Function ZWrkAdr_Cno%()
ZWrkAdr_Cno = ZZFndCno("WrkAdr")
End Function

Private Function ZZFndCno%(ColNm$)
Dim Ws As Worksheet
    Set Ws = ZCorr_Ws

Dim C%
    C = Ws_RC(Ws, 1, 1).End(xlToRight).Column

Dim Sqv
    Sqv = Ws_RCC(Ws, 1, 1, C).Value

Dim O%
    For O = 1 To UBound(Sqv, 2)
        If Sqv(1, O) = ColNm Then ZZFndCno = O: Exit Function
    Next
Stop
End Function
