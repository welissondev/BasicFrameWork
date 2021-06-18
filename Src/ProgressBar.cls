VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic.Components.ProgressBar")

Option Explicit

Private This As ProgressTypes
Private ProgressBar As MSForms.Label
Private PercentageCompleted As MSForms.Label
Private Container As MSForms.Frame
Private Percentage As Double

Private Type ProgressTypes
    Min As Long
    Max As Long
    Top As Variant
    Width As Variant
    Height As Variant
    Left As Variant
    BackColor As Variant
    BorderColor As Variant
    ContainerBackColor As Variant
End Type

Public Enum StandardProgressBarColor
    Comment = 10777186
    Cyan = 16640395
    Green = 6736896
    Orange = 7125247
    Pink = 13007359
    Purple = 16356285
    Red = 8612296
    Grey = &H80000010
    Blue = 12355149
End Enum

Public Enum StandardStyleProgressBar
    StFlat = 0
    StBorder = 1
    St3d = 2
    StEtched = 3
End Enum

Public Sub InitializeComponent(ByRef UI As Object, Optional ByVal ProgressBarBackColor As StandardProgressBarColor = Green, _
 Optional ProgressBarStyle As StandardStyleProgressBar = StFlat, Optional ByVal ContainerBackColor As StandardProgressBarColor = Grey)
    
    Dim ProgressBarControl As Control
    Dim PercentageCompletedControl As Control
    Dim ContainerControl As Control
    
    Set ContainerControl = CreateContainer(UI, ContainerBackColor, ProgressBarStyle)
    
    Set ProgressBarControl = CreateProgressBar(ContainerControl, ProgressBarBackColor, ProgressBarStyle)

    '//Adiciona PercentageCompleted
    Set PercentageCompleted = UI.Add("Forms.Label.1", "PercentageCompleted", True)
    Set PercentageCompletedControl = PercentageCompleted
    
    With PercentageCompletedControl
        .Top = ContainerControl.Top - 11
        .Left = ContainerControl.Left
        .Caption = "0%"
    End With

End Sub

Private Function CreateContainer(UI As Object, ContainerBackColor As StandardProgressBarColor, _
ProgressBarStyle As StandardStyleProgressBar) As Control
    
    Dim Control As Control
    
    Set Container = UI.Add("Forms.Frame.1", "Container", True)
    Set Control = Container
    
    '//Proporções
    With Control
        .Top = Top
        .Width = Width
        .Height = Height
        .Left = Left
    End With
    
    '//Cores
    With Control
        If (This.ContainerBackColor = Empty Or This.ContainerBackColor = Null) Then
            .BackColor = ContainerBackColor
        Else
            .BackColor = This.ContainerBackColor
        End If
    End With
    
    '//Styles
    With Control
        Select Case ProgressBarStyle
            Case Is = 0
                .SpecialEffect = StFlat
            Case Is = 1
                .BorderStyle = StBorder
                    If (This.BorderColor = Empty Or This.BorderColor = Null) Then
                        .BorderColor = BorderColor
                    Else
                        .BorderColor = This.BorderColor
                    End If
            Case Is = 2
                .SpecialEffect = St3d
            Case Is = 3
                .SpecialEffect = StEtched
        End Select
    End With
    
    Set CreateContainer = Container
    
End Function

Private Function CreateProgressBar(ContainerControl As Object, ProgressBarBackColor As StandardProgressBarColor, _
ProgressBarStyle As StandardStyleProgressBar) As Control
    
    Dim Control As Control
    
    Set ProgressBar = ContainerControl.Add("Forms.Label.1", "ProgressBar", True)
    Set Control = ProgressBar
    
    '//Proporções
    With Control
        .Width = 0
        .Height = Container.Height
        .Top = 0
        .Left = 0
    End With
    
    '//Cores
    With Control
        If (This.BackColor = Empty Or This.BackColor = Null) Then
            .BackColor = ProgressBarBackColor
        Else
            .BackColor = This.BackColor
        End If
        
    End With
    
    '//Styles
    With Control
        Select Case ProgressBarStyle
            Case Is = 0
                .SpecialEffect = 0
            Case Is = 1
                .BorderStyle = 0
            Case Is = 2
                .SpecialEffect = St3d
                .Top = -2
                .Left = -2
            Case Is = 3
                .BorderStyle = 0
        End Select
    End With
    
    Set CreateProgressBar = ProgressBar
    
End Function


Public Sub Refresh()

    Dim ProgressBarControl As Control
    Dim ContainerControl As Control
    
    Set ContainerControl = Container
    Set ProgressBarControl = ProgressBar
    
    Let Percentage = Min / Max
    Let PercentageCompleted = Format(Percentage, "0%") & " Concluido!"
    Let ProgressBarControl.Width = Percentage * (ContainerControl.Width)
    Let Min = Min + 1
    
End Sub

Public Sub Reset()
    
    ProgressBar.Width = 0
    PercentageCompleted.Caption = "0% Concluido"
    Percentage = 0
    
End Sub


Public Property Let Visible(Value As Boolean)
    Let Container.Visible = Value
    Let PercentageCompleted.Visible = Value
End Property



'//Cor de fundo do container
Public Property Get ContainerBackColor() As Variant
    If This.ContainerBackColor = Empty Or This.ContainerBackColor = Null Then
        Let ContainerBackColor = &H80000010
    Else
        Let ContainerBackColor = This.ContainerBackColor
    End If
End Property
Public Property Let ContainerBackColor(ByRef Value As Variant)
    If Value = Empty Or Value = Null Then
        This.ContainerBackColor = &H80000010
    Else
        This.ContainerBackColor = Value
    End If
End Property


'//Cor da barra de progresso
Public Property Get BackColor() As Variant
    If This.BackColor = Empty Or This.BackColor = Null Then
        Let BackColor = StandardProgressBarColor.Green
    Else
        Let BackColor = This.BackColor
    End If
End Property
Public Property Let BackColor(ByRef Value As Variant)
    If Value = Empty Or Value = Null Then
        This.BackColor = StandardProgressBarColor.Grey
    Else
        This.BackColor = Value
    End If
End Property

'//Cor da dorda de progresso
Public Property Get BorderColor() As Variant
    If This.BorderColor = Empty Or This.BorderColor = Null Then
        Let BorderColor = StandardProgressBarColor.Grey
    Else
        Let BorderColor = This.BorderColor
    End If
End Property
Public Property Let BorderColor(ByRef Value As Variant)
    If Value = Empty Or Value = Null Then
        This.BorderColor = StandardProgressBarColor.Grey
    Else
        This.BorderColor = Value
    End If
End Property


'//Valor minimo da barra de progresso
Public Property Get Min() As Long
    If This.Min <= 0 Then
        Min = 1
    Else
        Let Min = This.Min
    End If
End Property
Public Property Let Min(ByVal Value As Long)
    If Value <= 0 Then
        This.Min = 1
    Else
        This.Min = Value
    End If
End Property


'//Valor máximo da barra de progresso
Public Property Get Max() As Long
    Let Max = This.Max
End Property
Public Property Let Max(ByVal Value As Long)
    This.Max = Value
End Property


'//Distancia entre top do container e seu controle pai
Public Property Get Top() As Variant
    If This.Top = Empty Or This.Top = Null Then
        Let Top = 0
    Else
        Let Top = This.Top
    End If
End Property
Public Property Let Top(ByVal Value As Variant)
    If Value = Empty Or Value = Null Then
        This.Top = 0
    Else
        This.Top = Value
    End If
End Property


'//Distancia entre aesquerda do container e seu controle pai
Public Property Get Left() As Variant
    If This.Left = Empty Or This.Left = Null Then
        Let Left = 0
    Else
        Let Left = This.Left
    End If
End Property
Public Property Let Left(ByVal Value As Variant)
    If Value = Empty Or Value = Null Then
        This.Left = 0
    Else
        This.Left = Value
    End If
End Property


'//Tamanho do container
Public Property Get Width() As Variant
    If This.Width = Empty Or This.Width = Null Then
        Let Width = 270
    Else
        Let Width = This.Width
    End If
End Property
Public Property Let Width(ByVal Value As Variant)
    If Value = Empty Or Value = Null Then
        This.Width = 270
    Else
        This.Width = Value
    End If
End Property


'//Largura do container
Public Property Get Height() As Variant
    If This.Height = Empty Or This.Height = Null Then
        Let Height = 30
    Else
        Let Height = This.Height
    End If
End Property
Public Property Let Height(ByVal Value As Variant)
    If Value = Empty Or Value = Null Then
        This.Height = 30
    Else
        This.Height = Value
    End If
End Property


