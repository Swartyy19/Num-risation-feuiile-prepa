VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_2NVS 
   Caption         =   "UserForm1"
   ClientHeight    =   7392
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   14040
   OleObjectBlob   =   "UserForm_2NVS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_2NVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_2NVS()

End Sub
Private Sub btnVueA_Click()
    With ThisWorkbook.Sheets("Prépa Numérisée") ' À adapter si autre nom
        .Shapes("Image_2NVS_A").Visible = msoTrue
        .Shapes("Image_2NVS_B").Visible = msoFalse
    End With
    Unload Me
End Sub

Private Sub btnVueB_Click()
    With ThisWorkbook.Sheets("Prépa Numérisée")
        .Shapes("Image_2NVS_A").Visible = msoFalse
        .Shapes("Image_2NVS_B").Visible = msoTrue
    End With
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
