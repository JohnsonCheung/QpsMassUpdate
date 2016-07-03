Attribute VB_Name = "Xls_Ws_Crt_EvtMth0"
Option Explicit
Dim A_TarWs As Worksheet
Dim A_WsEvtNm$
Dim A_RfToPj$
Dim A_RfToMd$
Dim A_Fn$

Sub Tst_WsChg(Cell As Range)
MsgBox "Tst_WsChg"
End Sub

Sub Tst_WsSelChg(Cell As Range)
MsgBox "Tst_WsSelChg"
End Sub

Sub Ws_Clr_Md(Ws As Worksheet)
Set A_TarWs = Ws
Dim Md As CodeModule
    Set Md = ZTarMd
Md.DeleteLines 1, Md.CountOfLines
Md.InsertLines 1, "Option Explicit"
End Sub

Sub Ws_Crt_EvtMth_CallingFn(TarWs As Worksheet, WsEvtNm$, RfToPj$, Optional RfToMd$, Optional Fn$)
If IsNothing(TarWs) Then Exit Sub
Set A_TarWs = TarWs
A_WsEvtNm = WsEvtNm
A_RfToMd = RfToMd
A_RfToPj = RfToPj
A_Fn = Fn
'Stop
VBA.Interaction.DoEvents ' <- Try this if LstWs.CodeName will return something
ZDo_Add_WsEvtMeth_ToTarWs
ZDo_Add_WsRef_RfToPjMd
End Sub

Private Sub Tst()
Ws_Clr_Md yWsMassUpd
Ws_Crt_EvtMth_CallingFn yWsMassUpd, "Change", "jjMassUpd", , "Tst_WsChg"
Ws_Crt_EvtMth_CallingFn yWsMassUpd, "SelectionChange", "jjMassUpd", , "Tst_WsSelChg"
'Ws_Clr_Md yWsMassUpd
End Sub

Private Sub ZDo_Add_Ref()
ZTarPj.References.AddFromFile ZRfToPjFfn '<===
End Sub

Private Sub ZDo_Add_WsEvtMeth_ToTarWs()
ZTarMd.AddFromString ZWsEvtMthBdy
End Sub

Private Sub ZDo_Add_WsRef_RfToPjMd()
If ZIsTarPj_AlreadyHasRf Then Exit Sub
ZDo_Add_Ref
End Sub

Private Property Get ZIsTarPj_AlreadyHasRf() As Boolean
Dim Pj As VBProject
    Set Pj = ZTarPj
Dim A$
    A = ZRfToPjFfn
    
Dim O As Boolean
    If Pj.Filename = A Then
        O = True
    Else
        Dim Rf As Reference
        For Each Rf In Pj.References
            If Rf.FullPath = A Then O = True: Exit For  '<=== Already Referred, exit
        Next
    End If
ZIsTarPj_AlreadyHasRf = O
End Property

Private Function ZPjNm$(Pj As VBProject)
On Error Resume Next
ZPjNm = Pj.Name  ' If there is an Excel not yet saved, the Pj.Name will raise error
End Function

Private Property Get ZRfToPjFfn$()
Dim Pj As VBProject
For Each Pj In Application.VBE.VBProjects
    If ZPjNm(Pj) = A_RfToPj Then ZRfToPjFfn = Pj.Filename: Exit Property
Next
Er "{A_RfToPj} not in open", A_RfToPj
End Property

Private Property Get ZTarMd() As CodeModule
Dim WsCodeNm$
    WsCodeNm = A_TarWs.CodeName ' Using this code [A_LstWs.CodeName] will get empty string.  So try pass the WsCodeName as parameter
If WsCodeNm$ = "" Then Er "Given CodeName of given {Ws} is empty", A_TarWs.Name

Dim Cmp As VBComponent
For Each Cmp In ZTarPj.VBComponents
    If Cmp.Type = vbext_ct_Document Then
        If Cmp.Name = WsCodeNm Then Set ZTarMd = Cmp.CodeModule: Exit Property
    End If
Next
Er "Cannot find the CodeName of given {Ws}", A_TarWs.Name
End Property

Private Property Get ZTarPj() As VBProject
Dim A$
    A = ZTarWbFullNm
Dim Pj As VBProject
For Each Pj In Application.VBE.VBProjects
    If Pj.Filename = A Then Set ZTarPj = Pj: Exit Property
Next
End Property

Private Property Get ZTarWbFullNm$()
ZTarWbFullNm$ = Ws_Wb(A_TarWs).FullName
End Property

Private Property Get ZWsEvtMthBdy$()
Dim Fn$
    Fn = IIf(A_Fn = "", A_WsEvtNm, A_Fn)
Const L1$ = "Private Sub Worksheet_?(ByVal Target As Range)"
Const L2A$ = "?.? Target"
Const L2B$ = "?.?.? Target"
Const L3$ = "End Sub"
Const A = L1 & vbCrLf & L2A & vbCrLf & L3
Const B = L1 & vbCrLf & L2B & vbCrLf & L3
Dim O$
    If A_RfToMd = "" Then
        O = Fmt_QQ(A, A_WsEvtNm, A_RfToPj, Fn)
    Else
        O = Fmt_QQ(B, A_WsEvtNm, A_RfToPj, A_RfToMd, Fn)
    End If
ZWsEvtMthBdy = O
End Property
