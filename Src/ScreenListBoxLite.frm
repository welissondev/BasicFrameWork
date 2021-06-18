VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScreenListBoxLite 
   Caption         =   "Main-Screen"
   ClientHeight    =   8715.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17190
   OleObjectBlob   =   "ScreenListBoxLite.frx":0000
End
Attribute VB_Name = "ScreenListBoxLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Screen")

Private ListBox As ListBoxLite

Private Sub ButtonFillListbox_Click()
    
    ListBox.Row.Clear
    
    For I = 1 To TextBox1.Text
        With ListBox
            
            .Row.Add
            .Row.Cell.Item(I, 0) = Sheet.Cells(I, "A").Text
            .Row.Cell.Item(I, 1) = Sheet.Cells(I, "B").Text
            .Row.Cell.Item(I, 2) = Sheet.Cells(I, "C").Text
            .Row.Cell.Item(I, 3) = Sheet.Cells(I, "D").Text
            .Row.Cell.Item(I, 4) = Sheet.Cells(I, "M").Text
            .Row.Cell.Item(I, 5) = Sheet.Cells(I, "N").Text
            .Row.Cell.Item(I, 6) = Sheet.Cells(I, "O").Text
            .Row.Cell.Item(I, 7) = Sheet.Cells(I, "P").Text
            .Row.Cell.Item(I, 8) = Sheet.Cells(I, "Q").Text
            .Row.Cell.Item(I, 9) = Sheet.Cells(I, "R").Text
            .Row.Cell.Item(I, 10) = Sheet.Cells(I, "S").Text
            
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
