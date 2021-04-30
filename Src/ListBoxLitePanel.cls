VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxLitePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Top As Single
Public Left As Single
Public Height As Single
Public Width As Single
Public SpecialEffect As fmSpecialEffect
Public BorderStyle As fmBorderStyle
Public BorderColor As Variant
Public BackColor As Variant
Public Visible As Boolean
Public Tag As String
Public Container As MSForms.Frame

Private Sub Class_Initialize()
    Name = "Panel"
    Tag = "Panel"
    Top = 0
    Left = 0
    Height = 45
    Width = 145
    SpecialEffect = fmSpecialEffectFlat
    BorderStyle = fmBorderStyleSingle
    BorderColor = vbBlack
    Visible = True
End Sub
