VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Taraudage 
   Caption         =   "UserForm1"
   ClientHeight    =   5868
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10500
   OleObjectBlob   =   "UserForm_Taraudage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Taraudage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTaraudage1_Click()
    Call taraudage.AfficherTaraudage(1)
End Sub

Private Sub cmdTaraudage2_Click()
    Call taraudage.AfficherTaraudage(2)
End Sub

Private Sub cmdTaraudage3_Click()
    Call taraudage.AfficherTaraudage(3)
End Sub

Private Sub cmdSupprimer_Click()
    Dim estGauche As Boolean
    estGauche = CBool(Me.Tag) ' ? robustesse langue

    Dim niveau As Long
    niveau = CLng(Range("AP5").Value)

    Dim baseNom As String
    If estGauche Then
        baseNom = "Taraudage_V" & niveau & "_G"
    Else
        baseNom = "Taraudage_V" & niveau & "_D"
    End If

    Dim i As Long
    For i = 1 To 3
        On Error Resume Next
        Sheets("Prépa Numérisée").Shapes(baseNom & "_T" & i).Visible = msoFalse
        On Error GoTo 0
    Next i
    
    Unload UserForm_Taraudage
    MsgBox "Taraudages masqués pour le côté " & IIf(estGauche, "gauche", "droit") & ".", vbInformation
End Sub

Private Sub UserForm_Click()

End Sub
