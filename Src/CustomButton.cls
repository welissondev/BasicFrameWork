VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic.Components.CustomButton")

Public Font As NewFont

Public Event Click()
Public Event MouseMove()

Private WithEvents CustomButton As MSForms.CommandButton
Attribute CustomButton.VB_VarHelpID = -1
Private WithEvents Container As MSForms.Frame
Attribute Container.VB_VarHelpID = -1
Private Parent As Object

Private Type CustomButtonTypes
    Top As Single
    Width As Single
    Height As Single
    Left As Single
    BackColor As CustomButtonColors
    HouverBackColor As CustomButtonColors
    HouverBorderBackColor As CustomButtonColors
    HouverForeColor As CustomButtonColors
    BorderColor As CustomButtonColors
    ForeColor As CustomButtonColors
    BorderStyle As Byte
    SpecialEffect As Byte
    Parent As Object
    PictureBackGround As Object
    Tag As String
    Visible As Boolean
    ControlType As String
End Type

Public Enum CustomButtonColors
    BGPrimary = 16743168
    BGSecondary = 8222060
    BGSucess = 4564776
    BGDanger = 4535772
    BGWarning = 508415
    BGInfo = 12100119
    BGDark = 4209204
    BGWhite = 16777215
End Enum

Private This As CustomButtonTypes

Private Sub Class_Initialize()

    Set Font = New NewFont
    
    Font.Bold = True
    Font.Size = 10
    Font.Name = "Arial"
            
    This.BackColor = BGDark
    This.ForeColor = BGWhite
    This.HouverBackColor = BGSecondary
    This.HouverForeColor = BGWhite
    This.BorderColor = BGSecondary
    This.HouverBorderBackColor = BGDark
    This.BorderStyle = 1
    This.Visible = True

End Sub

Private Sub CustomButton_Click()
    RaiseEvent Click
End Sub

Private Sub CustomButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    
    Dim Ctl As Control
    
    For Each Ctl In Parent.Controls
        
        If TypeName(Ctl) = "Frame" And Ctl.Tag = "Container" Then
            
            If Ctl.BackColor <> This.BackColor Then
                Call ResetStyle
            End If
                       
        End If
        
    Next
    
    Container.BackColor = This.HouverBackColor
    Container.BorderColor = This.HouverBorderBackColor
    CustomButton.ForeColor = This.HouverForeColor
  
    RaiseEvent MouseMove
    
End Sub

Public Sub ResetStyle()
    
    Dim Ctl As Control
    
    For Each Ctl In Parent.Controls
        
        If TypeName(Ctl) = "Frame" And Ctl.Tag = "Container" Then
            
            If Ctl.BackColor <> This.BackColor Then
                Ctl.BackColor = This.BackColor
                Ctl.BorderColor = This.BorderColor
                Ctl.Controls.Item(CustomButton).ForeColor = This.ForeColor
            End If
                       
        End If
        
    Next
    
End Sub

Public Sub InitializeComponent(ObjectParent As Object, Name As String, Optional Caption As String)
    
    Dim Control As MSForms.Control

    Set Parent = ObjectParent
    Set Container = Parent.Controls.Add("Forms.Frame.1", "Container", True)
    Set Control = Container
    
    'Define proporções do container
    With Control
        .Height = Height
        .Width = Width
        .Top = Top
        .Left = Left
    End With
    
    'Define Propriedades do container
    With Container
        .BorderStyle = BorderStyle
        .BackColor = BackColor
        .BorderColor = BorderColor
        .SpecialEffect = SpecialEffect
        .Picture = PictureBackGround
        .PictureSizeMode = fmPictureSizeModeStretch
        .Tag = Tag
        .Visible = This.Visible
    End With
    
    'Adiciona botão no container
    Set CustomButton = Container.Controls.Add("Forms.CommandButton.1", Name, True)
    Set Control = CustomButton
    
    'Define Proporções do botão
    With Control
        .Width = Container.Width + 10
        .Height = Container.Height + 10
        .Top = .Top - 5
        .Left = .Left - 5
    End With
        
    'Define Propriedade
    With CustomButton
        Set .Font = Font
        .ForeColor = This.ForeColor
        .Caption = Caption
        .TakeFocusOnClick = False
        .BackStyle = fmBackStyleTransparent
    End With
        
End Sub

'Cor de fundo do container
Public Property Get PictureBackGround() As Object
    Set PictureBackGround = This.PictureBackGround
End Property

Public Property Let PictureBackGround(Value As Object)
    Set This.PictureBackGround = Value
End Property

'Cor de fundo do container
Public Property Get BackColor() As CustomButtonColors
    If This.BackColor = 0 Or This.BackColor = Null Then
        BackColor = CustomButtonColors.BGPrimary
    Else
        BackColor = This.BackColor
    End If
End Property

Public Property Let BackColor(Value As CustomButtonColors)
    If Value = 0 Or Value = Null Then
        This.BackColor = CustomButtonColors.BGPrimary
    Else
        This.BackColor = Value
    End If
End Property

'Cor de fundo ao pausar mouse
Public Property Get HouverBackColor() As CustomButtonColors
    HouverBackColor = This.HouverBackColor
End Property

Public Property Let HouverBackColor(Value As CustomButtonColors)
    If Value = 0 Then
        This.HouverBackColor = CustomButtonColors.BGSucess
    Else
        This.HouverBackColor = Value
    End If
End Property

'Cor da borda ao pausar mouse
Public Property Get HouverBorderBackColor() As CustomButtonColors
    HouverBorderBackColor = This.HouverBorderBackColor
End Property

Public Property Let HouverBorderBackColor(Value As CustomButtonColors)
    If Value = 0 Then
        This.HouverBorderBackColor = CustomButtonColors.BGDark
    Else
        This.HouverBorderBackColor = Value
    End If
End Property

'Cor da font do botão
Public Property Get ForeColor() As CustomButtonColors
    ForeColor = This.ForeColor
End Property

Public Property Let ForeColor(Value As CustomButtonColors)
    If Value = 0 Then
        This.ForeColor = CustomButtonColors.BGWhite
    Else
        This.ForeColor = Value
    End If
End Property

'Cor de fundo do texto ao pausar mouse
Public Property Get HouverForeColor() As CustomButtonColors
    HouverForeColor = This.HouverForeColor
End Property

Public Property Let HouverForeColor(Value As CustomButtonColors)
    If Value = 0 Then
        This.HouverForeColor = CustomButtonColors.BGPrimary
    Else
        This.HouverForeColor = Value
    End If
End Property

'Cor da borda do botão
Public Property Get BorderColor() As CustomButtonColors
    BorderColor = This.BorderColor
End Property

Public Property Let BorderColor(Value As CustomButtonColors)
    If Value = 0 Then
        This.BorderColor = CustomButtonColors.BGPrimary
    Else
        This.BorderColor = Value
    End If
End Property

'Estilo da borda do botão
Public Property Get BorderStyle() As Byte
    Let BorderStyle = This.BorderStyle
End Property

Public Property Let BorderStyle(Value As Byte)
    This.BorderStyle = Value
End Property

'Estilo da borda do botão
Public Property Get SpecialEffect() As Byte
    Let SpecialEffect = This.SpecialEffect
End Property

Public Property Let SpecialEffect(Value As Byte)
    This.SpecialEffect = Value
End Property

'Distancia entre top do container e seu controle pai
Public Property Get Top() As Single
    If This.Top = 0 Then
        Let Top = 0
    Else
        Let Top = This.Top
    End If
End Property

Public Property Let Top(Value As Single)
    If Value = 0 Then
        This.Top = 0
    Else
        This.Top = Value
    End If
End Property

'Distancia entre a esquerda do container e seu controle pai
Public Property Get Left() As Single
    If This.Left = 0 Then
        Left = 0
    Else
        Left = This.Left
    End If
End Property

Public Property Let Left(Value As Single)
    If Value = 0 Then
        This.Left = 0
    Else
        This.Left = Value
    End If
End Property

'Tamanho do container
Public Property Get Width() As Single
    If This.Width = 0 Then
        Width = 80
    Else
        Width = This.Width
    End If
End Property

Public Property Let Width(Value As Single)
    If Value = 0 Then
        This.Width = 80
    Else
        This.Width = Value
    End If
End Property

'Largura do container
Public Property Get Height() As Single
    If This.Height = 0 Then
        Height = 25
    Else
        Height = This.Height
    End If
End Property

Public Property Let Height(Value As Single)
    If Value = 0 Then
        This.Height = 25
    Else
        This.Height = Value
    End If
End Property

'Tag do controle
Public Property Get Tag() As String
    Tag = This.Tag
End Property

Public Property Let Tag(Value As String)
    This.Tag = Value
End Property

'Visibilidade do controle
Public Property Get Visible() As Boolean
    Visible = This.Visible
End Property

Public Property Let Visible(Value As Boolean)
    This.Visible = Value
    Container.Visible = Value
End Property

