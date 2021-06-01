VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxLiteEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic.Components.ListBoxLite")

Option Explicit

Private This As HandlerTypes

Private Type HandlerTypes
    Handler As ListBoxLiteHandler
End Type

Public WithEvents ListItem As MSForms.TextBox
Attribute ListItem.VB_VarHelpID = -1

Private Sub ListItem_Click()
    This.Handler.SetClick ListItem
End Sub

Private Sub ListItem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    This.Handler.SetMouseMove ListItem
End Sub

Public Property Set Handler(Value As ListBoxLiteHandler)
    Set This.Handler = Value
End Property

