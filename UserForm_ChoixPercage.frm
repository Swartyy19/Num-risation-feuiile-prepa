VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ChoixPercage 
   Caption         =   "UserForm1"
   ClientHeight    =   8088
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10128
   OleObjectBlob   =   "UserForm_ChoixPercage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_ChoixPercage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LigneID As Long
Private CôtéGauche As Boolean
Private Sub Label1_Click()

End Sub


Private Sub CommandButton1_Click()

End Sub

Private Sub UserForm_Click()

End Sub


Private Sub cmdVersDoubleBarre_Click()
    Me.Hide
    UserForm_Doublebarre.Show
End Sub

Private Sub AppliquerTypePercage(typePercage As String)
    Dim id As Integer
    Dim estGauche As Boolean
    Dim parts As Variant
    parts = Split(Me.Tag, "|")

    If UBound(parts) = 1 Then
        id = Val(parts(0))
        estGauche = CBool(parts(1))
    Else
        MsgBox "Erreur : Tag invalide.", vbCritical
        Exit Sub
    End If

    Dim niveau As Integer
    If IsNumeric(Sheets("Prépa Numérisée").Range("AP5").Value) Then
        niveau = CInt(Sheets("Prépa Numérisée").Range("AP5").Value)
    Else
        niveau = 1
    End If

    Dim suffixe As String
    suffixe = "_V" & niveau & "_" & IIf(estGauche, "G", "D") & id

    Dim formeNom As String
    Select Case Trim(typePercage)
        Case "Perçage face":     formeNom = "Percage_Face" & suffixe
        Case "Perçage latéral":  formeNom = "Percage_Lateral" & suffixe
        Case "CHC":              formeNom = "Percage_CHC" & suffixe
        Case "CHCF_PC":          formeNom = "CHCF_PC" & suffixe
        Case "CHCF":             formeNom = "CHCF" & suffixe
        Case "PF_PC":            formeNom = "PF_PC" & suffixe
        Case "CHC_CB":           formeNom = "CHC_CB" & suffixe
        Case "PF_CHCCH":         formeNom = "PF_CHCCH" & suffixe
        Case "PF_CHCCB":         formeNom = "PF_CHCCB" & suffixe
        Case "CHCA_PC":          formeNom = "CHCA_PC" & suffixe
        Case "CHCA":             formeNom = "CHCA" & suffixe
        Case "LG":               formeNom = "LG" & suffixe
        Case "LD":               formeNom = "LD" & suffixe
        Case Else
            MsgBox "Type de perçage inconnu : " & typePercage, vbCritical
            Exit Sub
    End Select

    ' Masquer tous les types de perçage pour cette ligne et ce côté uniquement
    Dim nom As Variant
    For Each nom In Array("Percage_Face", "Percage_Lateral", "Percage_CHC", _
                         "CHCF_PC", "CHCF", "PF_PC", "CHC_CB", "PF_CHCCH", _
                         "PF_CHCCB", "CHCA_PC", "CHCA", "LG", "LD")
        On Error Resume Next
        If estGauche Then
            Sheets("Prépa Numérisée").Shapes(nom & "_V" & niveau & "_G" & id).Visible = msoFalse
        Else
            Sheets("Prépa Numérisée").Shapes(nom & "_V" & niveau & "_D" & id).Visible = msoFalse
        End If
        On Error GoTo 0
    Next nom

    ' Affichage de la forme sélectionnée
    On Error Resume Next
    Sheets("Prépa Numérisée").Shapes(formeNom).Visible = msoTrue
    On Error GoTo 0

    ' Mise à jour de la cellule du type de perçage
    With Sheets("Prépa Numérisée")
        If estGauche Then
            .Range("AR" & id + 4).Value = typePercage
        Else
            .Range("AT" & id + 4).Value = typePercage
        End If
    End With

    ' Affichage automatique de la zone de commentaire spécifique
    Dim zoneNom As String
    zoneNom = "ZoneCommentaire_V" & niveau & "_" & IIf(estGauche, "G", "D") & id

    Dim grp As Shape
    Dim sh As Shape
    Dim zoneTexte As Shape
    On Error Resume Next
    Set grp = Sheets("Prépa Numérisée").Shapes(zoneNom)
    On Error GoTo 0

    If Not grp Is Nothing Then
        grp.Visible = msoTrue
        For Each sh In grp.GroupItems
            If sh.Type = msoTextBox Then
                Set zoneTexte = sh
                Exit For
            End If
        Next sh

        If Not zoneTexte Is Nothing Then
            Dim oldText As String
            On Error Resume Next
            oldText = zoneTexte.TextFrame.Characters.Text
            On Error GoTo 0

            Dim newText As String
            newText = InputBox("Ajouter un commentaire pour ce perçage :", "Commentaire", oldText)

            If Len(newText) > 0 Then
                zoneTexte.TextFrame.Characters.Text = newText
            End If
        Else
            MsgBox "Zone de texte introuvable dans le groupe '" & zoneNom & "'.", vbExclamation
        End If
    Else
        MsgBox "Zone de commentaire non trouvée : " & zoneNom, vbExclamation
    End If

    Unload Me
End Sub



Private Sub cmdFace_Click()
    AppliquerTypePercage "Perçage face"
End Sub

Private Sub cmdLateral_Click()
    AppliquerTypePercage "Perçage latéral"
End Sub

Private Sub cmdCHC_Click()
    AppliquerTypePercage "CHC"
End Sub

Private Sub cmdCHCF_PC_Click()
    AppliquerTypePercage "CHCF_PC"
End Sub

Private Sub cmdCHCF_Click()
    AppliquerTypePercage "CHCF"
End Sub

Private Sub cmdPF_PC_Click()
    AppliquerTypePercage "PF_PC"
End Sub

Private Sub cmdCHC_CB_Click()
    AppliquerTypePercage "CHC_CB"
End Sub

Private Sub cmdPF_CHCCH_Click()
    AppliquerTypePercage "PF_CHCCH"
End Sub

Private Sub cmdPF_CHCCB_Click()
    AppliquerTypePercage "PF_CHCCB"
End Sub

Private Sub cmdCHCA_PC_Click()
    AppliquerTypePercage "CHCA_PC"
End Sub

Private Sub cmdCHCA_Click()
    AppliquerTypePercage "CHCA"
End Sub

Private Sub cmdLG_Click()
    AppliquerTypePercage "LG"
End Sub

Private Sub cmdLD_Click()
    AppliquerTypePercage "LD"
End Sub

Private Sub cmdSupprimer_Click()
    Dim id As Long
    Dim estGauche As Boolean
    Dim donnees() As String
    Dim niveau As Long
    Dim suffixe As String
    Dim ws As Worksheet
    Set ws = Sheets("Prépa Numérisée")

    If InStr(Me.Tag, "|") = 0 Then
        MsgBox "Erreur : données incorrectes dans Tag", vbCritical
        Exit Sub
    End If

    donnees = Split(Me.Tag, "|")

    If IsNumeric(donnees(0)) Then
        id = CLng(donnees(0))
        estGauche = CBool(donnees(1)) ' ? CORRIGÉ ICI
    Else
        MsgBox "Erreur : données incorrectes dans Tag", vbCritical
        Exit Sub
    End If

    If IsNumeric(ws.Range("AP5").Value) Then
        niveau = CLng(ws.Range("AP5").Value)
    Else
        MsgBox "Erreur : Niveau invalide (AP5)", vbCritical
        Exit Sub
    End If

    suffixe = "_V" & niveau & "_" & IIf(estGauche, "G", "D") & id

    ' Masquer les formes de perçage du bon côté
    Dim nom As Variant
    For Each nom In Array("Percage_Face", "Percage_Lateral", "Percage_CHC", _
                          "CHCF_PC", "CHCF", "PF_PC", "CHC_CB", "PF_CHCCH", _
                          "PF_CHCCB", "CHCA_PC", "CHCA", "LG", "LD")
        On Error Resume Next
        ws.Shapes(nom & suffixe).Visible = msoFalse
        On Error GoTo 0
    Next nom

    ' Masquer la zone commentaire
    Dim zoneNom As String
    zoneNom = "ZoneCommentaire" & suffixe
    On Error Resume Next
    ws.Shapes(zoneNom).Visible = msoFalse
    On Error GoTo 0

    ' Réinitialiser les cellules associées
    If estGauche Then
        ws.Range("AR" & id + 4).Value = "Aucun"
    Else
        ws.Range("AT" & id + 4).Value = "Aucun"
    End If

    MsgBox "Perçage et commentaire masqués pour la ligne " & id, vbInformation


    Unload Me
End Sub



Public Sub Initialiser(id As Long, estGauche As Boolean)
    LigneID = id
    CôtéGauche = estGauche

    ' MAJ du label (ok)
    lblLigne.Caption = "Ligne de perçage : " & id & IIf(estGauche, " (Gauche)", " (Droite)")

    ' ?? AJOUT FONDAMENTAL : mise à jour du Tag
    Me.Tag = id & "|" & estGauche
End Sub


