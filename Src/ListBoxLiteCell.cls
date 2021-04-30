VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxLiteCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ListBox As ListBoxLite
Public BackColor As Variant
Public ForeColor As Variant
Public HoverForeColor As Variant
Public SpecialEffect As fmSpecialEffect
Public BorderStyle As fmBorderStyle
Public BorderColor As Variant

Private Type CellTypes
    Items As Object
End Type

Private This As CellTypes
Private WithEvents Handler As ListBoxLiteHandler
Attribute Handler.VB_VarHelpID = -1
Private Events As ListBoxLiteEvent
Private ArrayRowParseMouse As Object

Private Sub Class_Initialize()
    Set Handler = New ListBoxLiteHandler
    Set ArrayRowParseMouse = CreateObject("System.Collections.ArrayList")
    Set This.Items = CreateObject("System.Collections.ArrayList")
    Let BackColor = vbWhite
    Let ForeColor = vbBlack
    Let SpecialEffect = fmSpecialEffectFlat
    Let HoverForeColor = vbBlack
    Let BorderStyle = 1
    Let BorderColor = vbBlack
End Sub

Public Property Let Item(Row As Variant, Cell As Variant, Value As Variant)
    ListBox.Row.Items(Row - 1).Controls.Item(Cell).Text = Value
    DoEvents
End Property

Public Sub CreateCells(Row As Object)
    
    Dim Cell As Control
    Dim Object As Control
    Dim Count As Integer
    Dim Column As Object
        
    For Each Column In ListBox.Column.Group
        
        Let Count = This.Items.Count
        Set Cell = ListBox.AddComponent(TypeTextBoxCell, Row)
        
        With Cell
            .BackColor = BackColor
            .ForeColor = ForeColor
            .SpecialEffect = SpecialEffect
            .BorderStyle = BorderStyle
            .BorderColor = BorderColor
            .Width = Column.Width
            .Left = Column.Left
            .Height = 15
            .Top = Row.InsideHeight / 2 - .Height / 2
            .Visible = True
        End With
        
        Set Events = New ListBoxLiteEvent
        Set Events.Handler = Handler
        Set Events.ListItem = Cell
        
        This.Items.Add Events
        
    Next
    
End Sub

Public Property Get Items() As Object
    Set Items = This.Items
End Property

Private Sub Handler_MouseMove(Object As Object)
    
    On Error Resume Next
    
    Dim Row As Control
    Dim CellForeColor As Variant
        
    If Object.BackColor = vbGreen Then Exit Sub
    'If Object.ForeColor = CellForeColor Then Exit Sub
    
    CellForeColor = Object.ForeColor
    
    If ArrayRowParseMouse.Count > 0 Then
        For Each Row In ArrayRowParseMouse
            Row.BackColor = vbWhite
            Row.ForeColor = CellForeColor
            'Row.Parent.BackColor = Row.BackColor
        Next
        ArrayRowParseMouse.Clear
    End If
    
    Select Case Object.BackColor
        Case Is = vbWhite
            
            For Each Row In Object.Parent.Controls
                Row.BackColor = vbGreen
                Row.ForeColor = &H8000&
                'Row.Parent.BackColor = Row.BackColor
                ArrayRowParseMouse.Add Row
            Next
            
        Case Is = vbGreen
            For Each Row In Object.Parent.Controls
                Row.BackColor = vbWhite
                Row.ForeColor = vbBlack
                'Row.Parent.BackColor = Row.BackColor
                ArrayRowParseMouse.Add Row
            Next
    End Select
    
End Sub
