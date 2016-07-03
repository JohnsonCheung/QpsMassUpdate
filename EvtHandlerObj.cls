VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EvtHandlerObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A_Dic As New Dictionary

Sub Add_HandlerObj(Wb$, Ws$, Evt$, HandlerObj)
A_Dic.Add Wb & "." & Ws & "." & Evt, HandlerObj
End Sub

Function Get_HandlerObj(Wb$, Ws$, Evt$)
Get_HandlerObj = A_Dic(Wb & "." & Ws & "." & Evt)
End Function
