VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxLiteRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic.Components.ListBoxLite")

Option Explicit

Public Height As Single
Public TopMargin As Single
Public LeftMargin As Single
Public SpecialEffect As fmSpecialEffect
Public BorderStyle As fmBorderStyle
Public BorderColor As Variant
Public BackColor As Variant
Public ListBox As ListBoxLite
Public Cell As ListBoxLiteCell

Private Type RowTypes
    Items As Object
End Type

Private This As RowTypes

Private Sub Class_Initialize()
    Set This.Items = CreateObject("System.Collections.ArrayList")
    Let Height = 15
    Let TopMargin = -1
    Let LeftMargin = 0
    Let SpecialEffect = fmSpecialEffectFlat
    Let BorderStyle = 0
    Let BorderColor = vbBlack
    Let BackColor = vbWhite
End Sub

Public Sub Add()

    Dim Row As Control
    Set Row = ListBox.AddComponent(TypePanel, ListBox.Container)

    With Row
        
        .Tag = "PanelRowCell"
        .Height = Height
        .Left = LeftMargin
        .SpecialEffect = SpecialEffect
        .BorderStyle = BorderStyle
        .BorderColor = BorderColor
        .BackColor = BackColor
        .Font.Name = "Open Sans"

        If This.Items.Count > 0 Then
            .Top = This.Items.Item(This.Items.Count - 1).Top + (Height + TopMargin)
        Else
            .Top = .Top + ListBox.Column.Panel.Height + TopMargin
        End If
        
        .Width = ListBox.Column.Panel.Container.Width
        .Visible = True
        
    End With
    
    Call Items.Add(Row)
    Call Cell.CreateCells(Row)
    Let ListBox.ScrollBars = True
    
End Sub

Public Sub Clear()
 
    Dim Row As Control

    For Each Row In ListBox.Container.Controls
        If Row.Tag = "PanelRowCell" Then
            ListBox.Container.Controls.Remove (Row.Name)
        End If
        DoEvents
    Next
    
    Call This.Items.Clear
    Call Cell.Items.Clear
    Let ListBox.ScrollBars = False
 
End Sub

Public Sub AdjustSize(Index As Long)
    
    Dim Row As Control
     
    Let This.Items(Index).Width = 0
    
    For Each Row In ListBox.Column.Items
        This.Items(Index).Width = This.Items(Index).Width + Row.Width
    Next
    
End Sub

Public Property Get Items() As Object
    Set Items = This.Items
End Property
