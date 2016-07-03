Attribute VB_Name = "FmtFilter"
Option Explicit

Type YellowAdr
    C() As Integer
    R() As Long
End Type


Private A_Ws As Worksheet

Sub FmtFilter_DoRestore(WrkWs As Worksheet)
Set A_Ws = WrkWs
Dim C1%
    C1 = ZYellowCno
ZDo_Clr_Filler C1
ZDo_Restore_Color C1

Dim C2%
    C2 = ZRedCno
ZDo_Clr_Filler C2
ZDo_Restore_Color C2

End Sub

Sub FmtFilter_DoSetRed(WrkWs As Worksheet, Adr As RedAdr)
Set A_Ws = WrkWs
Dim Cno%
Cno = ZRedCno
ZDo_Set_Color_and_Filler Cno, ZRedRnoAy(Adr), rgbRed, rgbWhite
End Sub

Sub FmtFilter_DoSetYellow(WrkWs As Worksheet, Adr As YellowAdr)
Set A_Ws = WrkWs
ZDo_Set_Color_and_Filler ZYellowCno, ZYellowRnoAy(Adr), rgbYellow, rgbBlack
End Sub

Private Function ZDo_Clr_Filler(Cno%)
Dim R As Range
    Set R = ZRge(Cno)
R.ClearContents
R.VerticalAlignment = XlVAlign.xlVAlignTop
End Function

Private Sub ZDo_Restore_Color(Cno%)
With ZRge(Cno)
    .Interior.Color = 10092543  ' Light Yellow of filler color
    .Font.ColorIndex = xlAutomatic
End With
'== Restore header color ====
ZOrgWs_PivotTitRge.Copy
Ws_RC(A_Ws, 6, 1).PasteSpecial xlPasteFormats
End Sub

Private Sub ZDo_Set_Color_and_Filler(Cno%, RnoAy&(), Color&, FontColor&)
Dim J&
For J = 0 To UB(RnoAy)
    With Ws_RC(A_Ws, RnoAy(J), Cno)
        .Value = "X"
        .Interior.Color = Color
        .Font.Color = FontColor&
    End With
Next
End Sub

Private Function ZOrgWs() As Worksheet
Dim Wb As Workbook
Set Wb = A_Ws.Parent
Set ZOrgWs = Wb.Sheets(OrgWsNm)
End Function

Private Function ZOrgWs_PivotTitRge() As Range
Dim C&
C = A_Ws.ListObjects(1).DataBodyRange.Columns.Count
Set ZOrgWs_PivotTitRge = Ws_RCRC(ZOrgWs, 6, 1, 6, C) ' R1/R2=6 means title pivotable header
End Function

Private Property Get ZRedCno%()
'2nd ChrGp??Filler at row5
Dim C%
For C = ZYellowCno + 1 To A_Ws.Range("A5").End(xlToRight).Column
    If Ws_RC(A_Ws, 5, C) Like "ChrGp??Filler" Then ZRedCno = C: Exit Property
Next
End Property

Private Function ZRedRnoAy(Adr As RedAdr) As Long()
'Return an array of non-duplicated RnoAy by the redAdr, so that they will be put "X" as the filler
'The RedAdr.R must in the ListObjects(1)'s row range,
'    otherwise, this RedAdr.R will be skipped
Dim O&(), U&, J&, Ay$()
Dim R1&, R2& ' The Rno of the listobject(1)
Dim Red_Rno&  ' The Rno in each RedAdr

With Rge_R1R2(A_Ws.ListObjects(1).DataBodyRange)
    R1 = .R1
    R2 = .R2
End With

Ay = Adr.Adr
U = UB(Ay)
If U = -1 Then Exit Function
For J = 0 To U
    Red_Rno = CLng(Mid(Ay(J), Str_FirstDigitPos(Ay(J))))
    If R1 <= Red_Rno And Red_Rno <= R2 Then
        Push_NoDup O, Red_Rno
    End If
Next
ZRedRnoAy = O
End Function

Private Function ZRge(Cno%) As Range
Dim A As TR1R2
    A = Ws_ListObj_R1R2(A_Ws)
Dim R1&, R2&
    R1 = A.R1
    R2 = A.R2
Set ZRge = Ws_CRR(A_Ws, Cno, R1, R2)
End Function

Private Property Get ZYellowCno%()
'First ChrGp??Filler at row5
Dim C%
For C = 1 To A_Ws.Range("A5").End(xlToRight).Column
    If Ws_RC(A_Ws, 5, C) Like "ChrGp??Filler" Then ZYellowCno = C: Exit Property
Next
End Property

Private Function ZYellowRnoAy(Adr As YellowAdr) As Long()
Dim O&()
PushAy_NoDup O, Adr.R
ZYellowRnoAy = O
End Function
