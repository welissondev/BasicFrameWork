VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Frabasic")
Option Explicit

Public Function Generate(Optional Size As Integer = 8, Optional WithLetter As Boolean = False) As String
   
    Dim I As Integer
    Dim Letter As String
    Dim Number As Integer
    
    For I = 1 To Size
        
        Number = Int(9 * Rnd() & Int(7 * Rnd))
        Letter = Chr(Int((22 * Rnd()) + 1) + 64)
        
        Select Case WithLetter
            Case Is = False
                Generate = Generate & Number
                    
            Case Is = True
            
                If Number < 5 Then
                    Generate = Generate & Number & Letter
                Else
                    Generate = Generate & Letter & Number
                End If
                
        End Select
    Next
   
End Function

