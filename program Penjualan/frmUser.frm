VERSION 5.00
Begin VB.Form frmuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master User"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   5415
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   4080
         MouseIcon       =   "frmUser.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Bata&l"
         Height          =   495
         Left            =   1440
         MouseIcon       =   "frmUser.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "T&ambah"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmUser.frx":0A56
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   2760
         MouseIcon       =   "frmUser.frx":0D60
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   1440
         MouseIcon       =   "frmUser.frx":106A
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   4080
         MouseIcon       =   "frmUser.frx":1374
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBersih 
         Caption         =   "&Bersih"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmUser.frx":167E
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   2760
         MouseIcon       =   "frmUser.frx":1988
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.TextBox txtpass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtnmuser 
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtuserid 
      Height          =   285
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Master User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama User"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User Id"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsUser As New ADODB.Recordset
Dim sSQL As String
Dim sSQLInput As String

Private Sub cmdbatal_Click()
Call Reset
cmdCari.Enabled = True
End Sub

Private Sub cmdBersih_Click()
txtuserid.Text = Empty
txtnmuser.Text = Empty
txtpass.Text = Empty
cmdCari.Enabled = False
txtuserid.SetFocus
End Sub

Private Sub cmdCari_Click()
Dim strCari As String
Dim sSQLC As String
Dim RsCari As New ADODB.Recordset

strCari = InputBox("Masukkan User Id :", "Cari User")

sSQLC = "SELECT * FROM user WHERE userid='" & strCari & "'"
RsCari.Open sSQLC, Con, 3, 1

If strCari = Empty Then
    MsgBox "User Id masih kosong!", vbCritical, "Perhatian"
ElseIf RsCari.EOF Then
    MsgBox "Data tidak diketemukan!", vbExclamation, "Cari"
Else
    'MsgBox "Data ada!", vbInformation, "Ketemu"
    txtuserid.Text = RsCari!userid
    txtnmuser.Text = RsCari!nmuser
    txtpass.Text = RsCari!pass
    Call UnReset
    txtuserid.Enabled = False
    txtuserid.BackColor = &H80000013
    cmdEdit.Enabled = True
    cmdHapus.Enabled = True
    cmdBatal.Enabled = True
End If
End Sub

Private Sub cmdEdit_Click()
Dim sSQLEdit As String

txtuserid.Enabled = False
txtuserid.BackColor = &H80000013

If txtnmuser.Text = Empty Then
    MsgBox "Nama user masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtpass.Text = Empty Then
    MsgBox "Password masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
End If

sSQLEdit = "UPDATE user SET nmuser='" & txtnmuser.Text & "', pass='" & txtpass.Text & "' WHERE userid='" & txtuserid.Text & "'"
Con.Execute (sSQLEdit)
MsgBox "Data berhasil diedit", vbInformation, "Edit"
End Sub

Private Sub cmdHapus_Click()
Dim strHapusKode As String

strHapusKode = txtuserid.Text
If MsgBox("Yakin ingin menghapus user id : " & strHapusKode & " ?", vbQuestion + vbOKCancel, "Hapus User") = 1 Then
    Con.Execute ("DELETE FROM user WHERE userid='" & strHapusKode & "'")
    MsgBox "Data berhasil dihapus!", vbInformation, "Hapus"
    Call Reset
End If
End Sub

Private Sub cmdSimpan_Click()
Dim sUid As String

sUid = LTrim(txtuserid.Text)
sUid = RTrim(sUid)

If txtuserid.Text = Empty Then
    MsgBox "User Id masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtnmuser.Text = Empty Then
    MsgBox "Nama User masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtpass.Text = Empty Then
    MsgBox "Password masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
End If

sSQL = "SELECT * FROM user WHERE userid='" & sUid & "'"
RsUser.Open sSQL, Con, 3, 1

If RsUser.EOF Then
    sSQLInput = "INSERT INTO user(userid, nmuser, pass) VALUES('" & sUid & "', '" & txtnmuser.Text & "', '" & txtpass.Text & "')"
    'MsgBox (sSQLInput)
    MousePointer = 11
    Con.Execute (sSQLInput)
    MsgBox "Input data barang berhasil!", vbInformation, "Input"
    MousePointer = 0
    RsUser.Close
Else
    MsgBox "User Id sudah ada didatabase!", vbExclamation, "Perhatian!"
    RsUser.Close
    Exit Sub
End If

txtuserid.Enabled = False
txtnmuser.Enabled = False
txtpass.Enabled = False

txtuserid.BackColor = &H80000013
txtnmuser.BackColor = &H80000013
txtpass.BackColor = &H80000013

cmdBatal.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = True
cmdBersih.Enabled = False
cmdEdit.Enabled = True
cmdCari.Enabled = True
End Sub

Private Sub cmdTambah_Click()
cmdBatal.Enabled = True
cmdSimpan.Enabled = True
cmdBersih.Enabled = True
cmdCari.Enabled = False
cmdHapus.Enabled = False
cmdEdit.Enabled = False
Call UnReset
Call cmdBersih_Click
txtuserid.SetFocus
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call KonekDB
Call Reset
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdSimpan_Click
End If
End Sub

Private Sub txtuserid_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txtnmuser.SetFocus
End If
End Sub

Private Sub txtuserid_LostFocus()
txtuserid.Text = UCase(txtuserid.Text)
End Sub

Private Sub txtnmuser_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txtpass.SetFocus
End If
End Sub

Sub Reset()
txtuserid.Text = Empty
txtnmuser.Text = Empty
txtpass.Text = Empty

txtuserid.Enabled = False
txtnmuser.Enabled = False
txtpass.Enabled = False

txtuserid.BackColor = &H80000013
txtnmuser.BackColor = &H80000013
txtpass.BackColor = &H80000013

cmdBatal.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = False
cmdBersih.Enabled = False
cmdEdit.Enabled = False
End Sub

Sub UnReset()
txtuserid.Enabled = True
txtnmuser.Enabled = True
txtpass.Enabled = True

txtuserid.BackColor = &H80000005
txtnmuser.BackColor = &H80000005
txtpass.BackColor = &H80000005
End Sub



