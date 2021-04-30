VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainScreen 
   Caption         =   "Main-Screen"
   ClientHeight    =   8715.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17190
   OleObjectBlob   =   "MainScreen.frx":0000
End
Attribute VB_Name = "MainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ListBox As ListBoxLite

Private Sub ButtonFillListbox_Click()
    
    ListBox.Row.Clear
    
    For i = 1 To TextBox1.Text
        With ListBox
            
            .Row.Add
            .Row.Cell.Item(i, 0) = Sheet.Cells(i, "A").Text
            .Row.Cell.Item(i, 1) = Sheet.Cells(i, "B").Text
            .Row.Cell.Item(i, 2) = Sheet.Cells(i, "C").Text
            .Row.Cell.Item(i, 3) = Sheet.Cells(i, "D").Text
            .Row.Cell.Item(i, 4) = Sheet.Cells(i, "M").Text
            .Row.Cell.Item(i, 5) = Sheet.Cells(i, "N").Text
            .Row.Cell.Item(i, 6) = Sheet.Cells(i, "O").Text
            .Row.Cell.Item(i, 7) = Sheet.Cells(i, "P").Text
            .Row.Cell.Item(i, 8) = Sheet.Cells(i, "Q").Text
            .Row.Cell.Item(i, 9) = Sheet.Cells(i, "R").Text
            .Row.Cell.Item(i, 10) = Sheet.Cells(i, "S").Text
            
        End With
        
    Next
    
End Sub

Private Sub InitializeComponents()
    
    Set ListBox = New ListBoxLite
    
    With ListBox
        
        .Name = "ListBoxTest"
        .Panel.Left = 30
        .Panel.Top = 100
        .Panel.Height = 300
        .Panel.Width = 800
        .Parent = Me
        
        .Column.Name = "ColumnId"
        .Column.Text = "ID"
        .Column.Width = 40
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
           
        .Column.Name = "ColumnCode"
        .Column.Text = "Código"
        .Column.Width = 65
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnType"
        .Column.Text = "Tipo"
        .Column.Width = 40
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnCustomer"
        .Column.Text = "Cliente"
        .Column.Width = 150
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnCPF"
        .Column.Text = "CPF"
        .Column.Width = 75
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnRG"
        .Column.Text = "RG"
        .Column.Width = 65
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnCivilState"
        .Column.Text = "Estado Civil"
        .Column.Width = 68
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnFixePhone"
        .Column.Text = "Tefone"
        .Column.Width = 75
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnMobilePhone"
        .Column.Text = "Celular"
        .Column.Width = 80
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnWhatsApp"
        .Column.Text = "WhatsApp"
        .Column.Width = 80
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
        
        .Column.Name = "ColumnEmail"
        .Column.Text = "E-mail"
        .Column.Width = 400
        .Column.TypeColumn = TypeTextBoxColumn
        .Column.Add
               
    End With
    
End Sub

Private Sub UserForm_Initialize()
    InitializeComponents
End Sub
