Attribute VB_Name = "taraudage"
Public Sub AfficherTaraudage(typeNum As Integer)

    Dim estGauche As Boolean
    estGauche = CBool(UserForm_Taraudage.Tag)

    Dim niveauActuel As Long
    niveauActuel = CLng(Sheets("Prépa Numérisée").Range("AP5").Value)

    Dim prefix As String
    If estGauche Then
        prefix = "Taraudage_V"
    Else
        prefix = "Taraudage_V"
    End If

    ' Masquer tous les taraudages pour tous les niveaux (1 à 4) et tous les types (1 à 3)
    Dim v As Long, t As Long
    For v = 1 To 4
        For t = 1 To 3
            Dim nomForme As String
            If estGauche Then
                nomForme = prefix & v & "_G_T" & t
            Else
                nomForme = prefix & v & "_D_T" & t
            End If
            On Error Resume Next
            Sheets("Prépa Numérisée").Shapes(nomForme).Visible = msoFalse
            On Error GoTo 0
        Next t
    Next v

    ' Afficher uniquement la forme sélectionnée du bon niveau
    Dim baseNom As String
    If estGauche Then
        baseNom = "Taraudage_V" & niveauActuel & "_G"
    Else
        baseNom = "Taraudage_V" & niveauActuel & "_D"
    End If

    On Error Resume Next
    Sheets("Prépa Numérisée").Shapes(baseNom & "_T" & typeNum).Visible = msoTrue
    On Error GoTo 0
        ' ?? Fermer le UserForm une fois l’action faite
    Unload UserForm_Taraudage
    
End Sub

Public Sub AfficherUserFormTaraudage(estGauche As Boolean)
    With UserForm_Taraudage
        .Tag = CStr(estGauche)
        .Show
    End With
End Sub

