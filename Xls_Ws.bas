Attribute VB_Name = "Xls_Ws"
Option Explicit
Type TR1R2
    R1 As Long
    R2 As Long
End Type
Enum eOperator
    eOp_In = 1
    eOp_Fn = 2
End Enum

Sub R1R2_Dmp(P As TR1R2)
Debug.Print "R1R2=(" & P.R1 & "," & P.R2 & ")"
End Sub

Function Ws_A1_To_LastCell(Ws As Worksheet) As Range
Set Ws_A1_To_LastCell = Ws.Range(Ws.Cells(1, 1), Ws_LastCell(Ws))
End Function

Function Ws_Adr$(Ws As Worksheet, R, C, Optional IsAbsolute As Boolean)
Ws_Adr = Ws_RC(Ws, R, C).Address(IsAbsolute, IsAbsolute)
End Function

Sub Ws_AssertSingleListObj(Ws As Worksheet)
If Ws.ListObjects.Count = 1 Then Exit Sub
Const C = "{Worksheet} should 1 ListObject, but now it has {n} ListObject.  The {workbook} is {folder}"
Dim Wb As Workbook
Set Wb = Ws.Parent
Er C, Ws.Name, Ws.ListObjects.Count, Wb.Name, Wb_Pth(Wb)
End Sub

Function Ws_C(Ws As Worksheet, C) As Range
Set Ws_C = Ws_RC(Ws, 1, C).EntireColumn
End Function

Function Ws_CC(Ws As Worksheet, C1, C2) As Range
Set Ws_CC = Ws_RCC(Ws, 1, C1, C2).EntireColumn
End Function

Function Ws_CRR(Ws As Worksheet, C, R1, R2) As Range
Set Ws_CRR = Ws_RCRC(Ws, R1, C, R2, C)
End Function

Sub Ws_Chk_FldNmAy(Ws As Worksheet, FldNmAy$(), OEr$())
If IsNothing(Ws) Then Exit Sub
Dim C2%
    C2 = Ws.Range("A1").End(xlToRight).Column
Dim Sqv
    Sqv = Ws_RCC(Ws, 1, 1, C2).Value
Dim A$()
    Dim J%
    For J = 1 To UBound(Sqv, 2)
        Push A, Sqv(1, J)
    Next
    
Dim O$()
    For J = 0 To UB(FldNmAy)
        If Not Ay_Has(A, FldNmAy(J)) Then Push O, FldNmAy(J)
    Next
If Sz(O) = 0 Then Exit Sub
Push OEr, Fmt_QQ("Ws(?) of FldNm(?) has missed some fields:", Ws.Name, Join(Ay_Quote(Ay_StrEsc(A), "[]")))
PushAy OEr, Ay_AddPfx(O, "    ")
End Sub

Sub Ws_ClrNames(Ws As Worksheet, pPfx$)
Dim J%
Dim L%: L = Len(pPfx)
For J = Ws.Names.Count To 1 Step -1
    Dim iNm As Name: Set iNm = Ws.Names(J)
    Dim mNm$: mNm = iNm.Name
    If Left(mNm, L) = pPfx Then iNm.Delete
Next
End Sub

Sub Ws_CrtListObj(Ws As Worksheet)
Dim Cell1 As Range, Cell2 As Range, Rge As Range
Set Cell1 = Ws.Cells(1, 1)
Set Cell2 = Ws.Cells.SpecialCells(xlCellTypeLastCell)
Set Rge = Ws.Range(Cell1, Cell2)
Ws.ListObjects.Add xlSrcRange, Rge, , xlYes
End Sub

Sub Ws_CrtTbl_ByHdSqv_SrcSqv(Ws As Worksheet, HdSqv, SrcSqv)
Cell_PutSqv Ws.Range("A1"), HdSqv
Cell_PutSqv Ws.Range("A2"), SrcSqv
Ws_CrtListObj Ws
Dim C2%
    C2 = UBound(HdSqv, 2)
Ws.Columns("1:" & C2).AutoFit
End Sub

Sub Ws_Dlt(Ws As Worksheet)
Dim mXls As Application: Set mXls = Ws.Application
Dim mSave As Boolean: mSave = mXls.DisplayAlerts
mXls.DisplayAlerts = False
Ws.Delete
mXls.DisplayAlerts = mSave
End Sub

Function Ws_Dt(Ws As Worksheet) As TDt
Dim O_C2%:             O_C2 = Ws.Range("A1").End(xlToRight).Column
Dim R2&:                 R2 = Ws.Range("A1").End(xlDown).Row
Dim Rge As Range:   Set Rge = Ws_RCRC(Ws, 2, 1, R2, O_C2)
Dim Sqv:                Sqv = Rge.Value
Dim O_DrAy():        O_DrAy = Sqv_DrAy(Sqv)
Dim HdSqv:            HdSqv = Ws_RCC(Ws, 1, 1, O_C2).Value
Dim NFld%:             NFld = O_C2
Dim O_FldNmAy$(): O_FldNmAy = Sqv_Row1_ToAy(HdSqv, O_FldNmAy)
Dim O As TDt
    O.Nm = Ws.Name
    O.FldNmAy = O_FldNmAy
    O.DrAy = O_DrAy
Ws_Dt = O
End Function

Function Ws_Dt_Sel(Ws As Worksheet, WsFldNmAy$(), Optional AsFldNmLvs$, Optional WhereFldNm_UsingAsFldNm$, Optional Operator As eOperator, Optional Operand, Optional NewDtNm$) As TDt
Ws_Dt_Sel = Dt_Sel(Ws_Dt(Ws), WsFldNmAy, AsFldNmLvs, WhereFldNm_UsingAsFldNm, Operator, Operand, NewDtNm)
End Function

Sub Ws_FmtCol_AsTxt(Ws As Worksheet, C, R1&, R2&)
Ws_CRR(Ws, C, R1, R2).NumberFormat = "@"
End Sub

Function Ws_LastCell(Ws As Worksheet) As Range
Set Ws_LastCell = Ws.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function Ws_LastCol&(Ws As Worksheet)
Ws_LastCol& = Ws_LastCell(Ws).Column
End Function

Function Ws_LastCol_ByFirstListObj&(Ws As Worksheet)
With Ws.ListObjects(1).DataBodyRange
    Ws_LastCol_ByFirstListObj = .Column + .Columns.Count - 1
End With
End Function

Function Ws_LastRow&(Ws As Worksheet)
Ws_LastRow = Ws_LastCell(Ws).Row
End Function

Function Ws_ListObj_R1R2(Ws As Worksheet) As TR1R2
Dim R As Range
    Set R = Ws.ListObjects(1).DataBodyRange
Dim O As TR1R2
    O = Rge_R1R2(R)
Ws_ListObj_R1R2 = O
End Function

Function Ws_MaxCno&(Ws As Worksheet)
Ws_MaxCno = Ws.Cells.EntireColumn.Count
End Function

Function Ws_MaxRno&(Ws As Worksheet)
Ws_MaxRno = Ws.Cells.EntireRow.Count
End Function

Function Ws_New() As Worksheet
Dim O As Worksheet
Set O = Wb_New.Sheets(1)
If O.CodeName = "" Then Stop
Set Ws_New = O
End Function

Sub Ws_OutLine(Ws As Worksheet, Rno1%, Rno2%, Optional pLvl As Byte = 2)
Dim mRge As Range: Set mRge = Ws.Range(Ws.Cells(Rno1, 1), Ws.Cells(Rno2, 1))
mRge.EntireRow.OutlineLevel = pLvl
End Sub

Function Ws_R(Ws As Worksheet, R) As Range
Set Ws_R = Ws.Rows(R)
End Function

Function Ws_RC(Ws As Worksheet, R, C) As Range
Set Ws_RC = Ws.Cells(R, C)
End Function

Function Ws_RCC(Ws As Worksheet, R, C1, C2) As Range
Set Ws_RCC = Ws_RCRC(Ws, R, C1, R, C2)
End Function

Function Ws_RCRC(Ws As Worksheet, R1, C1, R2, C2) As Range
Dim Cell1 As Range, Cell2 As Range
Set Cell1 = Ws.Cells(R1, C1)
Set Cell2 = Ws.Cells(R2, C2)
Set Ws_RCRC = Ws.Range(Cell1, Cell2)
End Function

Function Ws_RR(Ws As Worksheet, R1, R2) As Range
Set Ws_RR = Ws_CRR(Ws, 1, R1, R2).EntireRow
End Function

Sub Ws_ShwAllDta(Ws As Worksheet)
Dim A As AutoFilter: Set A = Ws.AutoFilter
If TypeName(A) <> "Nothing" Then A.ShowAllData
End Sub

Sub Ws_Sort(Ws As Worksheet, pLvcCol$, Optional Rno As Byte = 1)
'Col in pLvcCol can have minus sign as prefix means descending
Ws_ShwAllDta Ws
Dim mA$(): mA = Split(pLvcCol, ",")
Dim mRnoEnd&: mRnoEnd = Ws.Range("A" & Rno).End(xlDown).Row
Dim mColEnd$: mColEnd = Chr(64 + Ws.Range("A" & Rno).End(xlToRight).Column)
Dim J%
With Ws.Sort
    With .SortFields
        .Clear
        For J = 0 To UBound(mA)
            Dim mAA$
            Dim mOrd As XlSortOrder
            If Right(mA(J), 1) = "-" Then
                mOrd = xlDescending
                mAA = Left(mA(J), Len(mA(J)) - 1)
            Else
                mOrd = xlAscending
                mAA = mA(J)
            End If
            Dim mAdr$: mAdr = mAA & Rno & ":" & mAA & mRnoEnd
            .Add Key:=Ws.Range(mAdr), Order:=mOrd
        Next
    End With
    mAdr = "A" & Rno & ":" & mColEnd & mRnoEnd
    .SetRange Ws.Range(mAdr)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .Apply                              '<== Sort the worksheet
End With
End Sub

Function Ws_Sqv(Ws As Worksheet)
If IsNothing(Ws) Then Exit Function
Dim LO As ListObject
    Set LO = Ws.ListObjects(1)
If IsNothing(LO) Then Exit Function
Dim Rge As Range
    Set Rge = LO.DataBodyRange
Ws_Sqv = Rge.Value
End Function

Function Ws_Sqv_ByA1DownRight_NoR1(Ws As Worksheet)
Dim R&, C&
R = Ws.Range("A1").End(xlDown).Row
C = Ws.Range("A1").End(xlToRight).Column
Ws_Sqv_ByA1DownRight_NoR1 = Ws_RCRC(Ws, 2, 1, R, C).Value
End Function

Function Ws_Sqv_ByA1ToLastCell_withR1(Ws As Worksheet)
'Find Sqv of From Cell-A1 to Last-Cell
Dim C1 As Range
Dim C2 As Range
Set C1 = Ws.Cells(1, 1)
Set C2 = Ws_LastCell(Ws)
Ws_Sqv_ByA1ToLastCell_withR1 = Ws.Range(C1, C2).Value
End Function

Function Ws_Wb(Ws As Worksheet) As Workbook
Set Ws_Wb = Ws.Parent
End Function

Sub Ws_Zoom(Ws As Worksheet, Zoom%)
Ws.Activate
ActiveWindow.Zoom = Zoom
End Sub

Private Sub Ws_Dt_SelFld__Tst()
Dim Dt As TDt
    Dt = Ws_Dt_Sel(ErWsV3, StrAy("CharName", "Must"), "ChrNm Must")
Stop
End Sub
