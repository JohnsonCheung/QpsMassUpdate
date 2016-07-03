Attribute VB_Name = "SrcWsInf"
Option Explicit
Type KeyDta
    Pj As String
    Sku As String
    QDte As Date
    KeyStr As String
    Rno As Long
End Type
Type TWsInf
    Ws As Worksheet
    HdSqv As Variant
    HdLinChr As Integer
    HdLinTit1 As Integer
    HdLinTit2 As Integer
    HdLinTit3 As Integer
    HdLinFldNm As Integer
    HdLinLbl As Integer
    Cno As TCno
    UR As Long
    UC As Long
    Sqv As Variant
    KeyDta() As KeyDta
    PkDic As New Dictionary
    CnoDef() As TCnoDef
End Type
Private A_Ws As Worksheet

Function TWsInf(Ws As Worksheet) As TWsInf
If IsNothing(Ws) Then Er "Given Ws is Nothing"
Set A_Ws = Ws
Dim NewFmt As Boolean
    If Ws.Range("A1") = "Watch Photo" Then
        NewFmt = False
    Else
        NewFmt = True
    End If

Dim O As TWsInf
    With O
        Set .Ws = Ws
        .HdSqv = ZHdSqv
        If NewFmt Then
            .HdLinChr = 1
            .HdLinFldNm = 2
            .HdLinTit1 = 3
            .HdLinTit2 = 4
            .HdLinTit3 = 5
        Else
            .HdLinTit1 = 1
            .HdLinTit2 = 2
            .HdLinTit3 = 3
            .HdLinChr = 4
            .HdLinFldNm = 5
        End If
        .HdLinLbl = 6
        .Sqv = Ws_Sqv(Ws)
        .Cno = TCno(.HdSqv, Ws.Name)
        .CnoDef = TCnoDef(.Cno)
        .KeyDta = ZKeyDta(.Sqv, .Cno.Key)
        Set .PkDic = ZPkDic(.Sqv, .Cno.Key)
        .UC = UBound(.Sqv, 2)
        .UR = UBound(.Sqv, 1)
    End With
TWsInf = O
End Function

Private Property Get ZHdSqv()
Dim Ws As Worksheet
    Set Ws = A_Ws
Dim LastCol&
    LastCol = Ws_LastCol_ByFirstListObj(Ws)

Dim Cell1 As Range
Dim Cell2 As Range
    Set Cell1 = Ws.Range("A1")
    Set Cell2 = Ws.Cells(6, LastCol)

Dim R As Range
    Set R = Ws.Range(Cell1, Cell2)

'The HdSqv row1-3 has merged cells, so, if a cell is empty, copy previous one.
Dim OSqv
    OSqv = R.Value
    Dim I%, J%
    For I = 1 To 3
        For J = 2 To UBound(OSqv, 2)
            If IsEmpty(OSqv(I, J)) Then OSqv(I, J) = OSqv(I, J - 1)
        Next
    Next
ZHdSqv = OSqv
End Property

Private Property Get ZKeyDta(Sqv, C As KeyCno) As KeyDta()
Dim UR&
    UR = UBound(Sqv, 1)
Dim R&
Dim O() As KeyDta
    ReDim O(1 To UR)
For R = 1 To UR
    O(R) = ZKeyDtaItm(Sqv, R, C)
Next
ZKeyDta = O
End Property

Private Property Get ZKeyDtaItm(Sqv, R&, C As KeyCno) As KeyDta
Dim O As KeyDta
    With O
        .Pj = Sqv(R, C.Pj)
        .Sku = Sqv(R, C.Sku)
        .QDte = Sqv(R, C.QDte)
        .Rno = R
        .KeyStr = Join(Array(.Pj, .Sku, .QDte), "|")
    End With
ZKeyDtaItm = O
End Property

Private Property Get ZPkDic(Sqv, C As KeyCno) As Dictionary
Set ZPkDic = ZPkDic_WithDupReturn(Sqv, C)
End Property

Private Function ZPkDic_WithDupReturn(Sqv, C As KeyCno, Optional ODupDic As Dictionary) As Dictionary
Dim SetDup As Boolean
    SetDup = Not IsNothing(ODupDic)

Dim O As New Dictionary
    O.CompareMode = TextCompare
    Dim J&
    For J = 1 To UBound(Sqv, 1)
        Dim A$(2)
            A(0) = Sqv(J, C.Pj)
            A(1) = Sqv(J, C.Sku)
            A(2) = Sqv(J, C.QDte)
        Dim K$
            K = Join(A, "|")
        
        If O.Exists(K) Then
            If SetDup Then ODupDic.Add K, O(K)
        Else
            O.Add K, J
        End If
    Next
Set ZPkDic_WithDupReturn = O
End Function
