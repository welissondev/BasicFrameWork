VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")

Option Explicit

'Propriedades referentes ao container
Public Name As String
Public Panel As New ListBoxLitePanel

'Propriedades referente a classes de composição
Public Column As ListBoxLiteColumn
Public Row As ListBoxLiteRow
Public Cell As ListBoxLiteCell

'Usado para implementar eventos referente ao container
Private WithEvents ListBox As MSForms.Frame
Attribute ListBox.VB_VarHelpID = -1

'Enúmeradores para definição de tipos do componente. Isso é últil para chamar
'o tipo de objeto que desejamos criar no container
Public Enum ListBoxComponentTypes
    TypePanel = 0
    TypeTextBoxColumn = 1
    TypeComboBoxColumn = 2
    TypeCheckBoxColumn = 3
    TypeButtonColumn = 4
    TypeImageColumn = 5
    TypeTextBoxCell = 6
End Enum

Private Type Propertys
    Parent As Object
End Type

Private This As Propertys

Private Sub Class_Initialize()
    
    With Panel
        .Top = 0
        .Left = 0
        .Width = 250
        .Height = 150
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .BorderColor = 1
        .BackColor = vbWhite
        .Visible = True
    End With
    
    Set Column = New ListBoxLiteColumn
    Set Row = New ListBoxLiteRow
    Set Cell = New ListBoxLiteCell
    
End Sub

Private Sub Initialize()
    
    'Nesse contexto o container é corpo da listbox. Ele é representado por um frame que agrupará
    'todos os components referentes ao controle. Ao passar o objeto <Parent> para a classe
    'esse método será chamado para criar o controle.
    Set ListBox = AddComponent(TypePanel, This.Parent)
    Set Panel.Container = ListBox
        
    With Panel.Container
        .Name = Name
        .Top = Panel.Top
        .Left = Panel.Left
        .Height = Panel.Height
        .Width = Panel.Width
        .SpecialEffect = Panel.SpecialEffect
        .BorderStyle = Panel.BorderStyle
        .BorderColor = Panel.BorderColor
        .BackColor = Panel.BackColor
        .Zoom = 130
        .Visible = Panel.Visible
    End With
    
    'Objetos relacionados a classe precisam acessar sua instancia, então elas
    'recebem a instancia do objeto que foi declarado no seu corpo para que possam
    'acessar a propriedades dessa classe
    Set Cell.ListBox = Me
    Set Column.ListBox = Me
    Set Row.ListBox = Me
    Set Row.Cell = Cell
        
End Sub

'Cria componente especificado no controle parente
Public Function AddComponent(IsType As ListBoxComponentTypes, ParentObject As Object) As Control

    Dim LsbTypeOfName, VbaTypeOfName As String
    Dim Component As Control

    Select Case IsType
        Case Is = TypePanel
            LsbTypeOfName = "Panel"
            VbaTypeOfName = "Forms.Frame.1"
                
        Case Is = TypeTextBoxColumn
            LsbTypeOfName = "TextBoxColumn"
            VbaTypeOfName = "Forms.TextBox.1"
        
        Case Is = TypeComboBoxColumn
            LsbTypeOfName = "ComboBoxColumn"
            VbaTypeOfName = "Forms.ComboBox.1"
        
        Case Is = TypeCheckBoxColumn
            LsbTypeOfName = "CheckBoxColumn"
            VbaTypeOfName = "Forms.CheckBox.1"
                
        Case Is = TypeButtonColumn
            LsbTypeOfName = "ButtonColumn"
            VbaTypeOfName = "Forms.Button.1"
        
        Case Is = TypeImageColumn
            LsbTypeOfName = "ImageColumn"
            VbaTypeOfName = "Forms.Image.1"
        
        Case Is = TypeTextBoxCell
            LsbTypeOfName = "TextBoxCell"
            VbaTypeOfName = "Forms.TextBox.1"
        
    End Select

    Set Component = ParentObject.Controls.Add(VbaTypeOfName, LsbTypeOfName, False)
    
    With Component
        .Tag = LsbTypeOfName
    End With
    
    Set AddComponent = Component

End Function

Public Sub SetScrolls()

    Dim RowCount As Long
    Dim InsideHeight As Single
    Dim InsideWidth As Single

    Let RowCount = Row.Items.Count
    
    With Container
        
        .ScrollTop = 0
        .ScrollLeft = 0

        Let InsideHeight = Row.Height * RowCount + (Row.TopMargin * RowCount) + (Row.Height + Row.TopMargin + Column.Panel.Container.Height) + 2
        Let InsideWidth = Column.Panel.Container.Width + 1
        
        If (.InsideHeight < InsideHeight And .InsideWidth < InsideWidth) Then
            .ScrollBars = fmScrollBarsBoth
            .ScrollHeight = InsideHeight
            .ScrollWidth = InsideWidth
            Exit Sub
        End If
        
        If (.InsideHeight < InsideHeight) Then
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = InsideHeight
            Exit Sub
        End If
            
        If (.InsideWidth < InsideWidth) Then
            .ScrollBars = fmScrollBarsHorizontal
            .ScrollWidth = InsideWidth
        End If
        
    End With

End Sub

Private Sub ListBox_AddControl(ByVal Control As MSForms.Control)
    If Control.Name = "PanelRowCell" Then
        Call SetScrolls
    End If
End Sub

Private Sub ListBox_RemoveControl(ByVal Control As MSForms.Control)
    If Control.Name = "PanelRowCell" Then
        Call SetScrolls
    End If
End Sub

Public Property Let Parent(Value As Object)
    Set This.Parent = Value: Call Initialize
End Property

Public Property Get Container() As Control
    Set Container = ListBox
End Property

Public Property Let ScrollBars(Value As Boolean)
    If Value = True Then
        SetScrolls
    Else
        ListBox.ScrollBars = fmScrollBarsNone
    End If
End Property

Public Property Get Object() As Object
    Set Object = Me
End Property

