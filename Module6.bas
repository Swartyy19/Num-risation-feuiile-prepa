Attribute VB_Name = "Module6"
Sub OuvreForme(id As Integer, estGauche As Boolean)
    With UserForm_ChoixPercage
        .Initialiser id, estGauche
        .Show
    End With
End Sub

