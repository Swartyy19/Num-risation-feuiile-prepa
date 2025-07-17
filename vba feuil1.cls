VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Bouton_TaraudageGauche_Click()
    Call AfficherUserFormTaraudage(True)
End Sub

Private Sub Bouton_Taraudage_Droite_Click()
    Call AfficherUserFormTaraudage(False)
End Sub

Public Sub AfficherUserFormTaraudage(estGauche As Boolean)
    With UserForm_Taraudage
        .Tag = CStr(estGauche)
        .Show
    End With
End Sub


Private Sub CommandButton1_Click()
Call AfficherUserFormTaraudage(True)
End Sub


Private Sub CommandButton2_Click()
    Call DupliquerFormesVisiblesFiltrées_V2
End Sub

Private Sub CommandButton3_Click()
    Call UniformiserHauteurFormes_V1_V2
End Sub

Private Sub CommandButton4_Click()
    Call CentrerFormes_V2_Cellules
End Sub

Private Sub ResetFormulaire_Click()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim colGauche As Variant, colDroite As Variant
    Dim colCotationGauche As Variant
    Dim col As Variant
    Dim i As Long

    Set ws = Sheets("Prépa Numérisée")
    Application.EnableEvents = False

    ' Réinitialiser les types de perçage (texte gauche/droite)
    For i = 5 To 10
        ws.Range("AR" & i).Value = "Aucun"  ' GAUCHE
        ws.Range("AT" & i).Value = "Aucun"  ' DROITE
    Next i

    ' Réinitialiser les cotations GAUCHE (colonnes de cotation V1 à V4 : AA, AL, AK, AI, AG, etc.)
    colCotationGauche = Array("AK", "AI", "AG", "AE", "AC", "AA")  ' adapte selon ton niveau max
    For Each col In colCotationGauche
        For i = 5 To 10
            On Error Resume Next
            ws.Range(col & i).MergeArea.ClearContents ' gère les cellules fusionnées
            On Error GoTo 0
        Next i
    Next col

    ' Réinitialiser les cotations visibles GAUCHE (ligne 6)
    colGauche = Array("L", "N", "P", "R", "T", "V")
    For i = 0 To UBound(colGauche)
        ws.Range(colGauche(i) & "6").ClearContents
    Next i

    ' Réinitialiser les cotations visibles DROITE (ligne 6)
    colDroite = Array("AF", "AH", "AB", "Z", "X", "Y")
    For i = 0 To UBound(colDroite)
        ws.Range(colDroite(i) & "6").ClearContents
    Next i

    ' Masquer TOUTES les formes de perçage visibles pour V1 et V2 (GAUCHE et DROITE)
    For Each shp In ws.Shapes
        If shp.Visible = msoTrue Then
            If shp.Name Like "*_V1_G#" Or shp.Name Like "*_V1_D#" Or _
               shp.Name Like "*_V2_G#" Or shp.Name Like "*_V2_D#" Then
                shp.Visible = msoFalse
            End If
        End If
    Next shp
    
    ' Masquer tous les taraudages (T1 à T3 pour V1 à V4, G et D)
    For v = 1 To 4
        For t = 1 To 3
            On Error Resume Next
            ws.Shapes("Taraudage_V" & v & "_G_T" & t).Visible = msoFalse
            ws.Shapes("Taraudage_V" & v & "_D_T" & t).Visible = msoFalse
            On Error GoTo 0
        Next t
    Next v

    MsgBox "Formulaire remis à zéro !", vbInformation
    Application.EnableEvents = True
End Sub





Private Sub Worksheet_Change(ByVal Target As Range)

    Dim id As Integer, nom As Variant
    Dim ligneDForme As Long, ligneGForme As Long
    Dim colD As Long, colG As Long
    Dim suffixeG As String, suffixeD As String
    Dim niveau As Integer
    Dim typePercage As String
    Dim formeNom As String

    ' ====== COTATION GAUCHE (colonne AS) ======
    If Not Intersect(Target, Range("AS5:AS20")) Is Nothing Then
        ligneGForme = Target.Row - 4
        colG = 12 + 2 * (ligneGForme - 1)
        Application.EnableEvents = False
        Cells(6, colG).Value = Target.Value
        Application.EnableEvents = True
    End If

    ' ====== COTATION DROITE (colonne AU) ======
    If Not Intersect(Target, Range("AU5:AU20")) Is Nothing Then
        ligneDForme = Target.Row - 4
        colD = 32 - 2 * (ligneDForme - 1)
        Application.EnableEvents = False
        Cells(6, colD).Value = Target.Value
        Application.EnableEvents = True
    End If

    ' ====== FORMES GAUCHE (AR5:AR20) ======
    If Not Intersect(Target, Range("AR5:AR20")) Is Nothing Then
        id = Target.Row - 4
        typePercage = Trim(Target.Value)
        If typePercage = "" Then Exit Sub

        If IsNumeric(Range("AP5").Value) Then
            niveau = CInt(Range("AP5").Value)
        Else
            niveau = 1
        End If

        suffixeG = "_V" & niveau & "_G" & id

        Select Case typePercage
            Case "CHCF_PC":     formeNom = "CHCF_PC" & suffixeG
            Case "CHCF":        formeNom = "CHCF" & suffixeG
            Case "PF_PC":       formeNom = "PF_PC" & suffixeG
            Case "CHC_CB":      formeNom = "CHC_CB" & suffixeG
            Case "PF_CHCCH":    formeNom = "PF_CHCCH" & suffixeG
            Case "PF_CHCCB":    formeNom = "PF_CHCCB" & suffixeG
            Case "CHCA_PC":     formeNom = "CHCA_PC" & suffixeG
            Case "CHCA":        formeNom = "CHCA" & suffixeG
            Case "LG":          formeNom = "LG" & suffixeG
            Case "LD":          formeNom = "LD" & suffixeG
            Case Else: Exit Sub
        End Select

        On Error Resume Next
        Shapes(formeNom).Visible = msoTrue
        On Error GoTo 0
    End If

    ' ====== FORMES DROITE (AT5:AT20) ======
    If Not Intersect(Target, Range("AT5:AT20")) Is Nothing Then
        id = Target.Row - 4
        typePercage = Trim(Target.Value)
        If typePercage = "" Then Exit Sub

        If IsNumeric(Range("AP5").Value) Then
            niveau = CInt(Range("AP5").Value)
        Else
            niveau = 1
        End If

        suffixeD = "_V" & niveau & "_D" & id

        Select Case typePercage
            Case "CHCF_PC":     formeNom = "CHCF_PC" & suffixeD
            Case "CHCF":        formeNom = "CHCF" & suffixeD
            Case "PF_PC":       formeNom = "PF_PC" & suffixeD
            Case "CHC_CB":      formeNom = "CHC_CB" & suffixeD
            Case "PF_CHCCH":    formeNom = "PF_CHCCH" & suffixeD
            Case "PF_CHCCB":    formeNom = "PF_CHCCB" & suffixeD
            Case "CHCA_PC":     formeNom = "CHCA_PC" & suffixeD
            Case "CHCA":        formeNom = "CHCA" & suffixeD
            Case "LG":          formeNom = "LG" & suffixeD
            Case "LD":          formeNom = "LD" & suffixeD
            Case Else: Exit Sub
        End Select

        On Error Resume Next
        Shapes(formeNom).Visible = msoTrue
        On Error GoTo 0
    End If

    ' ====== AFFICHAGE IMAGE PROFILÉ ======
    If Not Intersect(Target, Me.Range("AL7")) Is Nothing Then
        Dim profilé As String
        profilé = Trim(Target.Value)

        Dim profils As Variant
        profils = Array("30x30L", "40x40L", "45x45L", "40x80L", "45x90L", "80x80L", "90x90L")

        For Each nom In profils
            On Error Resume Next
            Me.Shapes("Image_" & nom).Visible = msoFalse
            On Error GoTo 0
        Next nom

        Me.Shapes("Image_2NVS_A").Visible = msoFalse
        Me.Shapes("Image_2NVS_B").Visible = msoFalse
        Me.Shapes("Profile_simple").Visible = msoFalse
        Me.Shapes("Profile_double").Visible = msoFalse
        Me.Shapes("Profile_quadruple").Visible = msoFalse

        If profilé <> "45x45_2NVS" Then
            On Error Resume Next
            Me.Shapes("Image_" & profilé).Visible = msoTrue
            On Error GoTo 0
        End If

        Select Case profilé
            Case "30x30L", "40x40L", "45x45L", "45x45_2NVS": Me.Shapes("Profile_simple").Visible = msoTrue
            Case "40x80L", "45x90L": Me.Shapes("Profile_double").Visible = msoTrue
            Case "80x80L", "90x90L": Me.Shapes("Profile_quadruple").Visible = msoTrue
        End Select
    End If

    ' ====== FORMULAIRE 2NVS ======
    If Target.Address = "$AL$7" Then
        If Trim(Target.Value) = "45x45_2NVS" Then
            UserForm_2NVS.Show
        End If
    End If

End Sub


Function ShapeExists(nom As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = Me.Shapes(nom)
    ShapeExists = Not shp Is Nothing
    On Error GoTo 0
End Function

' ----------- MODULE BOUTONS VERSION 1 -----------
' --- GAUCHE V1 ---
Sub Bouton_G1_Click(): AfficherFormulairePercageVx 1, True, "Bouton_G1": End Sub
Sub Bouton_G2_Click(): AfficherFormulairePercageVx 2, True, "Bouton_G2": End Sub
Sub Bouton_G3_Click(): AfficherFormulairePercageVx 3, True, "Bouton_G3": End Sub
Sub Bouton_G4_Click(): AfficherFormulairePercageVx 4, True, "Bouton_G4": End Sub
Sub Bouton_G5_Click(): AfficherFormulairePercageVx 5, True, "Bouton_G5": End Sub
Sub Bouton_G6_Click(): AfficherFormulairePercageVx 6, True, "Bouton_G6": End Sub

' --- DROITE V1 ---
Sub Bouton_D1_Click(): AfficherFormulairePercageVx 1, False, "Bouton_D1": End Sub
Sub Bouton_D2_Click(): AfficherFormulairePercageVx 2, False, "Bouton_D2": End Sub
Sub Bouton_D3_Click(): AfficherFormulairePercageVx 3, False, "Bouton_D3": End Sub
Sub Bouton_D4_Click(): AfficherFormulairePercageVx 4, False, "Bouton_D4": End Sub
Sub Bouton_D5_Click(): AfficherFormulairePercageVx 5, False, "Bouton_D5": End Sub
Sub Bouton_D6_Click(): AfficherFormulairePercageVx 6, False, "Bouton_D6": End Sub


' ----------- MODULE BOUTONS VERSION 2 -----------
' --- GAUCHE V2 ---
Sub BoutonV2_G1_Click(): AfficherFormulairePercageVx 1, True, "BoutonV2_G1": End Sub
Sub BoutonV2_G2_Click(): AfficherFormulairePercageVx 2, True, "BoutonV2_G2": End Sub
Sub BoutonV2_G3_Click(): AfficherFormulairePercageVx 3, True, "BoutonV2_G3": End Sub
Sub BoutonV2_G4_Click(): AfficherFormulairePercageVx 4, True, "BoutonV2_G4": End Sub
Sub BoutonV2_G5_Click(): AfficherFormulairePercageVx 5, True, "BoutonV2_G5": End Sub
Sub BoutonV2_G6_Click(): AfficherFormulairePercageVx 6, True, "BoutonV2_G6": End Sub

' --- DROITE V2 ---
Sub BoutonV2_D1_Click(): AfficherFormulairePercageVx 1, False, "BoutonV2_D1": End Sub
Sub BoutonV2_D2_Click(): AfficherFormulairePercageVx 2, False, "BoutonV2_D2": End Sub
Sub BoutonV2_D3_Click(): AfficherFormulairePercageVx 3, False, "BoutonV2_D3": End Sub
Sub BoutonV2_D4_Click(): AfficherFormulairePercageVx 4, False, "BoutonV2_D4": End Sub
Sub BoutonV2_D5_Click(): AfficherFormulairePercageVx 5, False, "BoutonV2_D5": End Sub
Sub BoutonV2_D6_Click(): AfficherFormulairePercageVx 6, False, "BoutonV2_D6": End Sub

