VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Doublebarre 
   Caption         =   "UserForm1"
   ClientHeight    =   7548
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   14652
   OleObjectBlob   =   "UserForm_Doublebarre.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Doublebarre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LigneID As Long
Private CôtéGauche As Boolean

Private Sub cmdPF_CHCCB_Click()
    AppliquerTypePercage "PF_CHCCB"
End Sub

Private Sub cmdCHCA_PC_Click()
    AppliquerTypePercage "CHCA_PC"
End Sub

Private Sub cmdCHCA_Click()
    AppliquerTypePercage "CHCA"
End Sub

Private Sub cmdLD_Click()
    AppliquerTypePercage "LD"
End Sub

Private Sub cmdCHCAB_PFH_Click()
    AppliquerTypePercage "CHCAB_PFH"
End Sub

Private Sub cmdCHCAB_PFH_PC_Click()
    AppliquerTypePercage "CHCAB_PFH_PC"
End Sub

Private Sub cmdCHCAH_CHCFB_Click()
    AppliquerTypePercage "CHCAH_CHCFB"
End Sub

Private Sub cmdCHCAH_CHCFB_PC_Click()
    AppliquerTypePercage "CHCAH_CHCFB_PC"
End Sub

Private Sub cmdCHCB_Click()
    AppliquerTypePercage "CHCB"
End Sub

Private Sub cmdCHCFB_Click()
    AppliquerTypePercage "CHCFB"
End Sub

Private Sub cmdCHCFB_PC_Click()
    AppliquerTypePercage "CHCFB_PC"
End Sub

Private Sub cmdCHCFB_PC__Click()
    AppliquerTypePercage "CHCFB_PC"
End Sub

Private Sub cmdCHCFH_CHCAB_Click()
    AppliquerTypePercage "CHCFH_CHCAB"
End Sub

Private Sub cmdCHCFH_CHCAB_PC_Click()
    AppliquerTypePercage "CHCFH_CHCAB_PC"
End Sub

Private Sub cmdCHCFH_CHCFB_Click()
    AppliquerTypePercage "CHCFH_CHCFB"
End Sub

Private Sub cmdCHCFH_CHCFB_PC_Click()
    AppliquerTypePercage "CHCFH_CHCFB_PC"
End Sub

Private Sub cmdCHCFH_Click()
    AppliquerTypePercage "CHCFH"
End Sub

Private Sub cmdCHCFH_PC__Click()
    AppliquerTypePercage "CHCFH_PC"
End Sub

Private Sub cmdCHCFH_PFB_Click()
    AppliquerTypePercage "CHCFH_PFB"
End Sub

Private Sub cmdCHCFH_PFB_PC_Click()
    AppliquerTypePercage "CHCFH_PFB_PC"
End Sub

Private Sub cmdCHCH_Click()
    AppliquerTypePercage "CHCH"
End Sub

Private Sub cmdPC_Click()
    AppliquerTypePercage "PC"
End Sub

Private Sub cmdPFB_CHCB_Click()
    AppliquerTypePercage "PFB_CHCB"
End Sub

Private Sub cmdPFB_CHCH_Click()
    AppliquerTypePercage "PFB_CHCH"
End Sub

Private Sub cmdPFB_Click()
    AppliquerTypePercage "PFB"
End Sub

Private Sub cmdPFB_PC_Click()
    AppliquerTypePercage "PFB_PC"
End Sub

Private Sub cmdPFH_CHCB_Click()
    AppliquerTypePercage "PFH_CHCB"
End Sub

Private Sub cmdPFH_CHCFB_Click()
    AppliquerTypePercage "PFH_CHCFB"
End Sub

Private Sub cmdPFH_CHCFB_PC_Click()
    AppliquerTypePercage "PFH_CHCFB_PC"
End Sub

Private Sub cmdPFH_CHCH_Click()
    AppliquerTypePercage "PFH_CHCH"
End Sub

Private Sub cmdPFH_Click()
    AppliquerTypePercage "PFH"
End Sub

Private Sub cmdCHCAB_Click()
    AppliquerTypePercage "CHCAB"
End Sub

Private Sub cmdCHCAB_PC_Click()
    AppliquerTypePercage "CHCAB_PC"
End Sub

Private Sub cmdCHCAH_Click()
    AppliquerTypePercage "CHCAH"
End Sub

Private Sub cmdCHCAH_CHCAB_Click()
    AppliquerTypePercage "CHCAH_CHCAB"
End Sub

Private Sub cmdCHCAH_CHCAB_PC_Click()
    AppliquerTypePercage "CHCAH_CHCAB_PC"
End Sub

Private Sub cmdCHCAH_PC_Click()
    AppliquerTypePercage "CHCAH_PC"
End Sub

Private Sub cmdCHCAH_PFB_Click()
    AppliquerTypePercage "CHCAH_PFB"
End Sub

Private Sub cmdCHCAH_PFB_PC_Click()
    AppliquerTypePercage "CHCAH_PFB_PC"
End Sub

Private Sub cmdPFH_PC_Click()
    AppliquerTypePercage "PFH_PC"
End Sub

Private Sub cmdPFH_PFB_CHCB_Click()
    AppliquerTypePercage "PFH_PFB_CHCB"
End Sub

Private Sub cmdPFH_PFB_CHCH_Click()
    AppliquerTypePercage "PFH_PFB_CHCH"
End Sub

Private Sub cmdPFH_PFB_Click()
    AppliquerTypePercage "PFH_PFB"
End Sub

Private Sub cmdPFH_PFB_PC_Click()
    AppliquerTypePercage "PFH_PFB_PC"
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
        Case "Perçage face":         formeNom = "Percage_Face" & suffixe
        Case "Perçage latéral":      formeNom = "Percage_Lateral" & suffixe
        Case "CHC":                  formeNom = "Percage_CHC" & suffixe
        Case "CHCF_PC":              formeNom = "CHCF_PC" & suffixe
        Case "CHCF":                 formeNom = "CHCF" & suffixe
        Case "PF_PC":                formeNom = "PF_PC" & suffixe
        Case "CHC_CB":               formeNom = "CHC_CB" & suffixe
        Case "PF_CHCCH":             formeNom = "PF_CHCCH" & suffixe
        Case "PF_CHCCB":             formeNom = "PF_CHCCB" & suffixe
        Case "CHCA_PC":              formeNom = "CHCA_PC" & suffixe
        Case "CHCA":                 formeNom = "CHCA" & suffixe
        Case "LG":                   formeNom = "LG" & suffixe
        Case "LD":                   formeNom = "LD" & suffixe
        Case "CHCAB":                formeNom = "CHCAB" & suffixe
        Case "CHCAB_PC":            formeNom = "CHCAB_PC" & suffixe
        Case "CHCAB_PFH":           formeNom = "CHCAB_PFH" & suffixe
        Case "CHCAB_PFH_PC":        formeNom = "CHCAB_PFH_PC" & suffixe
        Case "CHCAH":                formeNom = "CHCAH" & suffixe
        Case "CHCAH_CHCAB":         formeNom = "CHCAH_CHCAB" & suffixe
        Case "CHCAH_CHCFB":         formeNom = "CHCAH_CHCFB" & suffixe
        Case "CHCAH_CHCAB_PC":      formeNom = "CHCAH_CHCAB_PC" & suffixe
        Case "CHCAH_CHCFB_PC":      formeNom = "CHCAH_CHCFB_PC" & suffixe
        Case "CHCAH_PC":            formeNom = "CHCAH_PC" & suffixe
        Case "CHCAH_PFB":           formeNom = "CHCAH_PFB" & suffixe
        Case "CHCAH_PFB_PC":        formeNom = "CHCAH_PFB_PC" & suffixe
        Case "CHCFB":               formeNom = "CHCFB" & suffixe
        Case "CHCFB_PC":            formeNom = "CHCFB_PC" & suffixe
        Case "CHCFH":               formeNom = "CHCFH" & suffixe
        Case "CHCFH_CHCAB":         formeNom = "CHCFH_CHCAB" & suffixe
        Case "CHCFH_CHCAB_PC":      formeNom = "CHCFH_CHCAB_PC" & suffixe
        Case "CHCFH_CHCFB":         formeNom = "CHCFH_CHCFB" & suffixe
        Case "CHCFH_CHCFB_PC":      formeNom = "CHCFH_CHCFB_PC" & suffixe
        Case "CHCFH_PC":            formeNom = "CHCFH_PC" & suffixe
        Case "CHCFH_PFB":           formeNom = "CHCFH_PFB" & suffixe
        Case "CHCFH_PFB_PC":        formeNom = "CHCFH_PFB_PC" & suffixe
        Case "CHCH":                formeNom = "CHCH" & suffixe
        Case "CHCB":                formeNom = "CHCB" & suffixe
        Case "PC":                  formeNom = "PC" & suffixe
        Case "PFB":                 formeNom = "PFB" & suffixe
        Case "PFB_CHCB":            formeNom = "PFB_CHCB" & suffixe
        Case "PFB_CHCH":            formeNom = "PFB_CHCH" & suffixe
        Case "PFB_PC":              formeNom = "PFB_PC" & suffixe
        Case "PFH":                 formeNom = "PFH" & suffixe
        Case "PFH_CHCB":            formeNom = "PFH_CHCB" & suffixe
        Case "PFH_CHCFB":           formeNom = "PFH_CHCFB" & suffixe
        Case "PFH_CHCFB_PC":        formeNom = "PFH_CHCFB_PC" & suffixe
        Case "PFH_CHCH":            formeNom = "PFH_CHCH" & suffixe
        Case "PFH_PC":              formeNom = "PFH_PC" & suffixe
        Case "PFH_PFB":             formeNom = "PFH_PFB" & suffixe
        Case "PFH_PFB_CHCB":        formeNom = "PFH_PFB_CHCB" & suffixe
        Case "PFH_PFB_CHCH":        formeNom = "PFH_PFB_CHCH" & suffixe
        Case "PFH_PFB_PC":          formeNom = "PFH_PFB_PC" & suffixe
        Case Else
            MsgBox "Type de perçage inconnu : " & typePercage, vbCritical
            Exit Sub
    End Select

    ' Masquer tous les types de perçage pour cette ligne et ce côté uniquement
    Dim nom As Variant
    For Each nom In Array( _
    "Percage_Face", "Percage_Lateral", "Percage_CHC", "PFH_CHCH", "CHCF_PC", "CHCF", "PF_PC", "CHCAH_CHCFB", _
    "CHCAH_CHCFB", "CHC_CB", "PF_CHCCH", "PF_CHCCB", "CHCA_PC", "CHCA", "LG", "LD", "CHCAH_CHCFB_PC", _
    "CHCFB", "CHCFB_PC", "CHCFH", "CHCFH_PC", "CHCAB_PFH_PC", "CHCFH_CHCFB", "CHCFH_PFB", "CHCAB_PFH", "CHCFH_CHCFB_PC", _
    "CHCFH_CHCAB", "CHCFH_CHCAB_PC", "CHCH", "PC", "PFB", "PFH", "CHCAB", "CHCAB_PC", "CHCFH_PFB_PC", _
    "CHCAH", "CHCB", "CHCAH_PC", "PFH_PC", "PFH_PFB", "PFH_PFB_PC", "CHCAH_PFB", "CHCFH_CHCAB_PC", "CHCAH_PFB_PC", _
    "CHCAH_CHCAB", "CHCAH_CHCAB_PC", "PFH_CHCB", "PFH_CHCFB", "PFH_CHCFB_PC", "PFH_CHCH", "PFB_CHCB", _
    "PFH_PFB_CHCB", "PFH_PFB_CHCH", "PFB_CHCB", "PFB_CHCH", "PFB_CHCH", "PFB_PC" _
)



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
        estGauche = CBool(donnees(1))
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

    ' ? suffixe correctement formé, sans redondance
    suffixe = "_V" & niveau & "_" & IIf(estGauche, "G", "D") & id

    ' ? Liste propre des types de perçage (pas de _G ou _D ici)
    Dim nom As Variant
    For Each nom In Array("Percage_Face", "Percage_Lateral", "Percage_CHC", _
    "CHCF_PC", "CHCF", "PF_PC", "CHC_CB", "PF_CHCCH", "PF_CHCCB", _
    "CHCA_PC", "CHCA", "LG", "LD", "CHCB", "CHCAH_CHCFB_PC", _
    "CHCAB", "CHCAB_PC", "CHCAB_PFH", "PFH_CHCH", "CHCAB_PFH_PC", _
    "CHCAH", "CHCAH_PC", "CHCAH_PFB", "CHCFH_PFB", "CHCFH_CHCAB", "CHCAH_PFB_PC", _
    "CHCAH_CHCAB", "CHCAH_CHCAB_PC", "CHCAH_CHCFB", "CHCAH_CHCFB_PC", _
    "CHCFB", "CHCFB_PC", "CHCFH", "CHCFH_PC", "CHCFH_PFB_PC", _
    "CHCFH_CHCAB", "CHCFH_CHCAB_PC", "CHCFH_CHCFB", "CHCFH_CHCFB_PC", _
    "CHCH", "PC", "PFB", "PFB_PC", "PFB_CHCB", "PFB_CHCH", _
    "PFH", "PFH_PC", "PFH_PFB", "PFH_PFB_PC", "CHCAH_CHCFB", "PFB_CHCB", _
    "PFH_CHCB", "PFH_CHCFB", "PFH_CHCFB_PC", "PFH_CHCH", "PFB_CHCH", _
    "PFH_PFB_CHCB", "PFH_PFB_CHCH", "CHCAH_CHCFB")
                         
        On Error Resume Next
        ws.Shapes(nom & suffixe).Visible = msoFalse
        On Error GoTo 0
    Next nom

    ' ? Zone commentaire
    Dim zoneNom As String
    zoneNom = "ZoneCommentaire" & suffixe
    On Error Resume Next
    ws.Shapes(zoneNom).Visible = msoFalse
    On Error GoTo 0

    ' ? Mise à jour des cellules
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

    ' MAJ du label si tu en as un, sinon commente cette ligne
    ' lblLigne.Caption = "Ligne de perçage : " & id & IIf(estGauche, " (Gauche)", " (Droite)")

    ' Mise à jour du Tag pour passage d’infos
    Me.Tag = id & "|" & estGauche
End Sub




Private Sub UserForm_Click()

End Sub

Private Sub cmdRetourChoixPerçage_Click()
    Me.Hide
    UserForm_ChoixPercage.Show
End Sub

