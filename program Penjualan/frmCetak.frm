VERSION 5.00
Begin VB.Form frmCetak 
   Caption         =   "Cetak kartu Anggota"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "frmCetak.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCetak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub
