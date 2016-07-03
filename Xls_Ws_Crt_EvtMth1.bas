Attribute VB_Name = "Xls_Ws_Crt_EvtMth1"
Option Explicit
Private A_WsEvtLvsNm$
Private A_TarWs As Worksheet
Private A_RfToPjNm$

Sub Ws_Crt_EvtMth_CallingObjFn(TarWs As Worksheet, WsEvtLvsNm$, HandlerObj, Optional RfToPjNm$)
If IsNothing(TarWs) Then Er "Given TarWs is nothing"
Set A_TarWs = TarWs
A_RfToPjNm = RfToPjNm
A_WsEvtLvsNm = WsEvtLvsNm
VBA.Interaction.DoEvents ' <- Try this if LstWs.CodeName will return something
ZDo_Add_WsEvtMeth_ToTarWs
ZDo_Add_WsRef_RfToPjMd
ZDo_Add_HandlerObj HandlerObj
End Sub

Private Sub Tst()
'Dim M As New Macro_ErWs_V3SelInErWs
'Dim Handler
'    Set Handler = M.Init(yWsMassUpd)
'Ws_Clr_Md yWsMassUpd
'Ws_Crt_EvtMth_CallingObjFn yWsMassUpd, "Change SelectionChange", Handler
'Debug.Print EvtHandlerObjDic.Count
''Ws_Clr_Md yWsMassUpd
'XX = 3
End Sub

Private Property Get ZBdy_WsEvtMthBdy$()
Dim A$()
    A = ZEvtNmAy
Dim O$()
    Dim J%
    For J = 0 To UB(A)
        Push O, ZBdy_WsEvtMthBdy_OneEvt(A(J))
    Next
ZBdy_WsEvtMthBdy = Join(O, vbCrLf & vbCrLf)
End Property

Private Property Get ZBdy_WsEvtMthBdy_OneEvt$(EvtNm$)
Const L1$ = "Friend Sub Worksheet_?(ByVal Cell As Range)"
Const L2$ = "Dim Obj"
Const L3$ = "Set Obj = EvtHandlerObjDic(""[?][?][?]"")"
Const L4$ = "CallByName Obj,""?"",VbMethod,Cell"
Const L5$ = "End Sub"
Const A = L1 & vbCrLf & L2 & vbCrLf & L3 & vbCrLf & L4 & vbCrLf & L5
ZBdy_WsEvtMthBdy_OneEvt = Fmt_QQ(A, EvtNm, ZWbNm, ZWsNm, EvtNm, EvtNm)
End Property

Private Sub ZDo_Add_HandlerObj(HandlerObj)
Dim A$()
    A = ZEvtNmAy

Dim Wb$
    Wb = ZWbNm
Dim Ws$
    Ws = ZWsNm
Dim J%
For J = 0 To UB(A)
    Dim K$
        K = Fmt_QQ("[?][?][?]", Wb, Ws, A(J))
    Debug.Print K
    EvtHandlerObjDic.Add K, HandlerObj   '<== Add
Next
Debug.Print EvtHandlerObjDic.Count
End Sub

Private Sub ZDo_Add_Ref()
ZTarPj.References.AddFromFile ZRfToPjFfn '<===
End Sub

Private Sub ZDo_Add_WsEvtMeth_ToTarWs()
ZTarMd.AddFromString ZBdy_WsEvtMthBdy
End Sub

Private Sub ZDo_Add_WsRef_RfToPjMd()
If A_RfToPjNm = "" Then Exit Sub
If ZIsTarPj_AlreadyHasRf Then Exit Sub
ZDo_Add_Ref
End Sub

Private Property Get ZEvtNmAy() As String()
ZEvtNmAy = Ay_RmvBlankEle(Split(A_WsEvtLvsNm))
End Property

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
If A_RfToPjNm = "" Then Exit Property
Dim Pj As VBProject
For Each Pj In Application.VBE.VBProjects
    If ZPjNm(Pj) = A_RfToPjNm Then ZRfToPjFfn = Pj.Filename: Exit Property
Next
Er "{A_RfToPj} not in open", A_RfToPjNm
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

Private Property Get ZWbNm$()
ZWbNm = Ws_Wb(A_TarWs).Name
End Property

Private Property Get ZWsNm$()
ZWsNm = A_TarWs.Name
End Property
