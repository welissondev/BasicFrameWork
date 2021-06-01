VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxLiteColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic.Components.ListBoxLite")
Option Explicit
Option Compare Binary

'Ajusta margem entre a borda direita da última coluna
'e a borda direita do painel
Private Const ADJUSTMENT_MARGIN As Single = 1.5

Public Name As String
Public Text As String
Public Width As Single
Public Height As Single
Public Margin As Single
Public ForeColor As Variant
Public TypeColumn As ListBoxComponentTypes
Public ListBox As ListBoxLite
Public Panel As ListBoxLitePanel

Private Type ColumnTypes
    Group As Object
End Type

Private This As ColumnTypes

Private Sub Class_Initialize()
    
    Set Panel = New ListBoxLitePanel
    Set This.Group = CreateObject("System.Collections.ArrayList")
    
    Let Height = 15
    Let Margin = -1
    Let ForeColor = vbWhite
    
    With Panel
        .Name = "ColumnPanel"
        .Tag = .Name
        .Top = 0
        .Left = 0
        .Height = 15
        .Width = 0
        .BorderStyle = 0
        .BorderColor = vbRed
        .BackColor = vbBlack
        .SpecialEffect = 0
        .Visible = True
    End With
    
End Sub

Public Sub Add()
    
    Select Case TypeColumn
        Case Is = TypeTextBoxColumn: AddTypeTextColumn
    End Select
    
End Sub

Private Sub AddTypeTextColumn()

    Dim NextColumn As Control
    Dim LastColumn As Control
    Dim Count As Integer
    
    Let Count = Group.Count
        
    Call CreatePanel
        
    'Adiciona a colunas no <Panel> criado para agrupar colunas
    Set NextColumn = ListBox.AddComponent(TypeTextBoxColumn, Panel.Container)
    With NextColumn
    
        .Name = Name
        .Text = Text
        .Width = Width + Margin
        .Height = Height
        .ForeColor = ForeColor
        .BackColor = Panel.Container.BackColor
        .SpecialEffect = 0
        .BorderStyle = 0
        .BorderColor = vbBlue
        .Visible = True
             
        'Posiciona as colunas de acordo com a coluna anterior
        If Count > 0 Then
            Set LastColumn = Group(Count - 1)
            Let NextColumn.Left = LastColumn.Left + (LastColumn.Width) + Margin
        End If
        
        'Centraliza a coluna no centro do seu painel
        .Top = Panel.Container.InsideHeight / 2 - .Height / 2
        
    End With
    
    'Adiciona objeto coluna <TextBox> no <ArrayList>
    Group.Add [NextColumn]
    
    Call AdjustPanelSize
    Call ListBox.SetScrolls
    
End Sub

'Ajusta o tamnho do painel de acordo com a quantidade de colunas
Private Sub AdjustPanelSize()
    
    Dim Size As Single
    Dim Column As Control
    Dim Count As Integer
    
    Count = Panel.Container.Controls.Count

    For Each Column In Panel.Container.Controls
        Size = Size + Column.Width
    Next

    Panel.Container.Width = (Margin * Count + Margin) + Size + ADJUSTMENT_MARGIN
    
End Sub

'Adiciona o <Panel> para agrupar colunas
Private Sub CreatePanel()
    
    If (Panel.Container Is Nothing) Then
        Set Panel.Container = ListBox.AddComponent(TypePanel, ListBox.Container)
        With Panel.Container
            .Name = Panel.Name
            .Tag = Panel.Tag
            .Top = Panel.Top
            .Left = Panel.Left
            .Height = Panel.Height
            .Width = Panel.Width
            .BorderStyle = Panel.BorderStyle
            .BorderColor = Panel.BorderColor
            .BackColor = Panel.BackColor
            .SpecialEffect = Panel.SpecialEffect
            .Visible = Panel.Visible
        End With
    End If
    
End Sub

Public Property Get Group() As Object
    Set Group = This.Group
End Property

