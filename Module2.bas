Attribute VB_Name = "Module2"
Sub Bouton_Taraudage_Gauche_Click()
    AfficherUserFormTaraudage True ' Gauche
End Sub



Sub AfficherUserFormTaraudage(estGauche As Boolean)
    UserForm_Taraudage.Tag = estGauche
    UserForm_Taraudage.Show
End Sub

Sub Bouton_Taraudage_Droite_Click()
    AfficherUserFormTaraudage False ' Droite
End Sub

