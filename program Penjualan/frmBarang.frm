VERSION 5.00
Begin VB.Form frmbrg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Barang"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   5415
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   4080
         MouseIcon       =   "frmBarang.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Bata&l"
         Height          =   495
         Left            =   1440
         MouseIcon       =   "frmBarang.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "T&ambah"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmBarang.frx":0A56
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   2760
         MouseIcon       =   "frmBarang.frx":0D60
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   1440
         MouseIcon       =   "frmBarang.frx":106A
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   4080
         MouseIcon       =   "frmBarang.frx":1374
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBersih 
         Caption         =   "&Bersih"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmBarang.frx":167E
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   2760
         MouseIcon       =   "frmBarang.frx":1988
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.ComboBox cbSatuan 
      Height          =   315
      ItemData        =   "frmBarang.frx":1C92
      Left            =   1320
      List            =   "frmBarang.frx":1CA2
      TabIndex        =   3
      Text            =   "---------------Pilih---------------"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtstok 
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txthrgbrg 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtnmbrg 
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtkdbrg 
      Height          =   285
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Master Barang"
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
      TabIndex        =   10
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label5 
      Caption         =   "Stok"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Satuan"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Harga Barang"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Barang"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Barang"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmbrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsBrg As New ADODB.Recordset
Dim sSQL As String
Dim sSQLInput As String

'Private Sub cbSatuan_Click()
'txtstok.SetFocus
'End Sub

Private Sub cbSatuan_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txtstok.SetFocus
End If
End Sub

Private Sub cmdbatal_Click()
Call Reset
cmdCari.Enabled = True
End Sub

Private Sub cmdBersih_Click()
txtkdbrg.Text = Empty
txtnmbrg.Text = Empty
txthrgbrg.Text = Empty
cbSatuan.Text = "---------------Pilih---------------"
txtstok.Text = Empty
cmdCari.Enabled = False
txtkdbrg.SetFocus
End Sub

Private Sub cmdCari_Click()
Dim strCari As String
Dim sSQLC As String
Dim RsCari As New ADODB.Recordset

strCari = InputBox("Masukkan Kode Barang :", "Cari Barang")

sSQLC = "SELECT * FROM barang WHERE kdbrg='" & strCari & "'"
RsCari.Open sSQLC, Con, 3, 1

If strCari = Empty Then
    MsgBox "Kode barang masih kosong!", vbCritical, "Perhatian"
ElseIf RsCari.EOF Then
    MsgBox "Data tidak diketemukan!", vbExclamation, "Cari"
Else
    'MsgBox "Data ada!", vbInformation, "Ketemu"
    txtkdbrg.Text = RsCari!kdbrg
    txtnmbrg.Text = RsCari!nmbrg
    txthrgbrg.Text = RsCari!hrgbrg
    cbSatuan.Text = RsCari!satuan
    txtstok.Text = RsCari!stok
    Call UnReset
    txtkdbrg.Enabled = False
    txtkdbrg.BackColor = &H80000013
    cmdEdit.Enabled = True
    cmdHapus.Enabled = True
    cmdBatal.Enabled = True
End If
End Sub

Private Sub cmdEdit_Click()
Dim sSQLEdit As String

txtkdbrg.Enabled = False
txtkdbrg.BackColor = &H80000013

If txtnmbrg.Text = Empty Then
    MsgBox "Nama barang masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txthrgbrg.Text = Empty Then
    MsgBox "Harga masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf cbSatuan.Text = Empty Or cbSatuan.Text = "---------------Pilih---------------" Then
    MsgBox "Satuan masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtstok.Text = Empty Then
    MsgBox "Stok masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
End If

sSQLEdit = "UPDATE barang SET nmbrg='" & txtnmbrg.Text & "', hrgbrg='" & CDbl(txthrgbrg.Text) & "', satuan='" & cbSatuan.Text & "', stok='" & CInt(txtstok.Text) & "' WHERE kdbrg='" & txtkdbrg.Text & "'"
Con.Execute (sSQLEdit)
MsgBox "Data berhasil diedit", vbInformation, "Edit"
End Sub

Private Sub cmdHapus_Click()
Dim strHapusKode As String

strHapusKode = txtkdbrg.Text
If MsgBox("Yakin ingin menghapus kode barang : " & strHapusKode & " ?", vbQuestion + vbOKCancel, "Hapus Barang") = 1 Then
    Con.Execute ("DELETE FROM barang WHERE kdbrg='" & strHapusKode & "'")
    MsgBox "Data berhasil dihapus!", vbInformation, "Hapus"
    Call Reset
End If
End Sub

Private Sub cmdSimpan_Click()
Dim sBrgn As String

sBrgn = LTrim(txtkdbrg.Text)
sBrgn = RTrim(sBrgn)

If txtkdbrg.Text = Empty Then
    MsgBox "Kode barang masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtnmbrg.Text = Empty Then
    MsgBox "Nama barang masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txthrgbrg.Text = Empty Then
    MsgBox "Harga masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf cbSatuan.Text = Empty Or cbSatuan.Text = "---------------Pilih---------------" Then
    MsgBox "Satuan masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtstok.Text = Empty Then
    MsgBox "Stok masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
End If

sSQL = "SELECT * FROM barang WHERE kdbrg='" & sBrgn & "'"
RsBrg.Open sSQL, Con, 3, 1

If RsBrg.EOF Then
    sSQLInput = "INSERT INTO barang(kdbrg, nmbrg, hrgbrg, satuan, stok) VALUES('" & sBrgn & "', '" & txtnmbrg.Text & "', '" & CDbl(txthrgbrg.Text) & "', '" & cbSatuan.Text & "', '" & CInt(txtstok.Text) & "')"
    'MsgBox (sSQLInput)
    MousePointer = 11
    Con.Execute (sSQLInput)
    MsgBox "Input data barang berhasil!", vbInformation, "Input"
    MousePointer = 0
    RsBrg.Close
Else
    MsgBox "Kode Barang sudah ada didatabase!", vbExclamation, "Perhatian!"
    RsBrg.Close
    Exit Sub
End If

txtkdbrg.Enabled = False
txtnmbrg.Enabled = False
txthrgbrg.Enabled = False
cbSatuan.Enabled = False
txtstok.Enabled = False

txtkdbrg.BackColor = &H80000013
txtnmbrg.BackColor = &H80000013
txthrgbrg.BackColor = &H80000013
cbSatuan.BackColor = &H80000013
txtstok.BackColor = &H80000013

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
txtkdbrg.SetFocus
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call KonekDB
Call Reset
End Sub

Private Sub txthrgbrg_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(";") Or KeyAscii = vbKeySpace Or KeyAscii = 13) Then
    txthrgbrg.Text = Empty
    MsgBox "Harap masukkan harga dengan angka!", vbCritical, "Perhatian!"
    Beep
    KeyAscii = 0
    Exit Sub
ElseIf KeyAscii = 13 Then
    cbSatuan.SetFocus
End If
End Sub

Private Sub txtkdbrg_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txtnmbrg.SetFocus
End If
End Sub

Private Sub txtkdbrg_LostFocus()
txtkdbrg.Text = UCase(txtkdbrg.Text)
End Sub

Private Sub txtnmbrg_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txthrgbrg.SetFocus
End If
End Sub

Private Sub txtstok_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(";") Or KeyAscii = vbKeySpace Or KeyAscii = 13) Then
    txthrgbrg.Text = Empty
    MsgBox "Harap masukkan jumlah stok dengan angka!", vbCritical, "Perhatian!"
    Beep
    KeyAscii = 0
    Exit Sub
ElseIf KeyAscii = 13 Then
    Call cmdSimpan_Click
End If
End Sub

Sub Reset()
txtkdbrg.Text = Empty
txtnmbrg.Text = Empty
txthrgbrg.Text = Empty
cbSatuan.Text = "---------------Pilih---------------"
txtstok.Text = Empty

txtkdbrg.Enabled = False
txtnmbrg.Enabled = False
txthrgbrg.Enabled = False
cbSatuan.Enabled = False
txtstok.Enabled = False

txtkdbrg.BackColor = &H80000013
txtnmbrg.BackColor = &H80000013
txthrgbrg.BackColor = &H80000013
cbSatuan.BackColor = &H80000013
txtstok.BackColor = &H80000013

cmdBatal.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = False
cmdBersih.Enabled = False
cmdEdit.Enabled = False
End Sub

Sub UnReset()
txtkdbrg.Enabled = True
txtnmbrg.Enabled = True
txthrgbrg.Enabled = True
cbSatuan.Enabled = True
txtstok.Enabled = True

txtkdbrg.BackColor = &H80000005
txtnmbrg.BackColor = &H80000005
txthrgbrg.BackColor = &H80000005
cbSatuan.BackColor = &H80000005
txtstok.BackColor = &H80000005
End Sub

