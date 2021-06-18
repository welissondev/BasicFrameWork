VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxLiteHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic.Components.ListBoxLite")

Option Explicit

Event Click(Object As Object)
Event MouseMove(Object As Object)

Public Sub SetClick(Sender As Object)
    RaiseEvent Click(Sender)
End Sub

Public Sub SetMouseMove(Sender As Object)
    RaiseEvent MouseMove(Sender)
End Sub
