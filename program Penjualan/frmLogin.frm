VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   2400
      MouseIcon       =   "frmLogin.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "&Login"
      Height          =   495
      Left            =   600
      MouseIcon       =   "frmLogin.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtuserid 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Password    :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "User Id        :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbatal_Click()
Unload Me
End Sub

Private Sub cmdLogin_Click()
Dim RsLogin As New ADODB.Recordset
Dim sSQL As String

sSQL = "SELECT * FROM user WHERE userid='" & txtuserid.Text & "'"
RsLogin.Open sSQL, Con, 3, 1

If RsLogin.EOF Then
    MsgBox "User Id belum ada didatabase!", vbExclamation, "Perhatian"
    Exit Sub
ElseIf txtpass.Text = RsLogin!pass Then
    MenuUtama.Show
    Unload Me
Else
    MsgBox "Password salah!", vbCritical, "Login"
End If
End Sub

Private Sub Form_Load()
Call KonekDB
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdLogin_Click
End If
End Sub

Private Sub txtuserid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtpass.SetFocus
End If
End Sub

Private Sub txtuserid_LostFocus()
txtuserid.Text = UCase(txtuserid.Text)
End Sub
