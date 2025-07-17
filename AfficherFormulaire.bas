Attribute VB_Name = "AfficherFormulaire"
Sub AfficherFormulairePercageVx(ligne As Long, estGauche As Boolean, nomBouton As String)
    Dim typeProfile As String
    Dim celluleType As Range
    Dim niveau As Integer

    ' Détection automatique du niveau selon le nom du bouton
    If InStr(nomBouton, "V2") > 0 Then
        niveau = 2
    Else
        niveau = 1
    End If
    Sheets("Prépa Numérisée").Range("AP5").Value = niveau ' Mise à jour du niveau actif

    ' Détection du profil sélectionné
    Set celluleType = Sheets("Prépa Numérisée").Range("AL7")
    typeProfile = Trim(celluleType.Value)

    ' Affichage du bon formulaire selon le type de profil
    Select Case typeProfile
        Case "30x30L", "40x40L", "45x45L", "45x45_2NVS"
            With UserForm_ChoixPercage
                .Tag = ligne & "|" & estGauche
                .Show
            End With
        Case "45x90L", "40x80L"
            With UserForm_Doublebarre
                .Tag = ligne & "|" & estGauche
                .Show
            End With
        Case Else
            MsgBox "Type de profilé inconnu : " & typeProfile, vbExclamation
    End Select
End Sub

