VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmcust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Customer"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmcustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrintKA 
      Caption         =   "&Cetak Kartu Anggota"
      Height          =   1095
      Left            =   3720
      MouseIcon       =   "frmcustomer.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   1920
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtTglMsk 
      Height          =   375
      Left            =   1320
      TabIndex        =   21
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   93454337
      CurrentDate     =   41617
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   5415
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   4080
         MouseIcon       =   "frmcustomer.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Bata&l"
         Height          =   495
         Left            =   1440
         MouseIcon       =   "frmcustomer.frx":0A56
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "T&ambah"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmcustomer.frx":0D60
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   2760
         MouseIcon       =   "frmcustomer.frx":106A
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   1440
         MouseIcon       =   "frmcustomer.frx":1374
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   4080
         MouseIcon       =   "frmcustomer.frx":167E
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBersih 
         Caption         =   "&Bersih"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmcustomer.frx":1988
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   2760
         MouseIcon       =   "frmcustomer.frx":1C92
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.ComboBox cbjenkel 
      Height          =   315
      ItemData        =   "frmcustomer.frx":1F9C
      Left            =   1320
      List            =   "frmcustomer.frx":1FA6
      TabIndex        =   3
      Text            =   "---------------Pilih---------------"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txthp 
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtalamat 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtnmcust 
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtkdcust 
      Height          =   285
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "bln/tgl/tahun"
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Tgl Masuk"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Master Customer"
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
      Caption         =   "No HP"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Customer"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Customer"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmcust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCust As New ADODB.Recordset
Dim sSQL As String
Dim sSQLInput As String

Private Sub cbjenkel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txthp.SetFocus
End If
End Sub

Private Sub cmdbatal_Click()
Call Reset
cmdCari.Enabled = True
End Sub

Private Sub cmdBersih_Click()
txtkdcust.Text = Empty
txtnmcust.Text = Empty
txtalamat.Text = Empty
cbjenkel.Text = "---------------Pilih---------------"
txthp.Text = Empty
'dtTglMsk.Value = Empty
cmdCari.Enabled = False
txtkdcust.SetFocus
End Sub

Private Sub cmdCari_Click()
Dim strCari As String
Dim sSQLC As String
Dim RsCari As New ADODB.Recordset

strCari = InputBox("Masukkan kode customer :", "Cari Customer")

sSQLC = "SELECT * FROM customer WHERE kdcust='" & strCari & "'"
RsCari.Open sSQLC, Con, 3, 1

If strCari = Empty Then
    MsgBox "Kode customer masih kosong!", vbCritical, "Perhatian"
ElseIf RsCari.EOF Then
    MsgBox "Data tidak diketemukan!", vbExclamation, "Cari"
Else
    'MsgBox "Data ada!", vbInformation, "Ketemu"
    txtkdcust.Text = RsCari!kdcust
    txtnmcust.Text = RsCari!nmcust
    txtalamat.Text = RsCari!alamat
    cbjenkel.Text = RsCari!jenkel
    txthp.Text = RsCari!nohp
    dtTglMsk.Value = RsCari!tglmsk
    Call UnReset
    txtkdcust.Enabled = False
    txtkdcust.BackColor = &H80000013
    cmdEdit.Enabled = True
    cmdHapus.Enabled = True
    cmdBatal.Enabled = True
End If
End Sub

Private Sub cmdEdit_Click()
Dim sSQLEdit As String

txtkdcust.Enabled = False
txtkdcust.BackColor = &H80000013

If txtnmcust.Text = Empty Then
    MsgBox "Nama customer masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtalamat.Text = Empty Then
    MsgBox "Alamat masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf cbjenkel.Text = Empty Or cbjenkel.Text = "---------------Pilih---------------" Then
    MsgBox "Jenis kelamin masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txthp.Text = Empty Then
    MsgBox "Nomor HP masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf dtTglMsk.Value = Empty Then
    MsgBox "Tanggal masuk masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
End If

sSQLEdit = "UPDATE customer SET nmcust='" & txtnmcust.Text & "', alamat='" & txtalamat.Text & "', jenkel='" & cbjenkel.Text & "', nohp='" & CStr(txthp.Text) & "', tglmsk='" & dtTglMsk.Value & "' WHERE kdcust='" & txtkdcust.Text & "'"
Con.Execute (sSQLEdit)
MsgBox "Data berhasil diedit", vbInformation, "Edit"
End Sub

Private Sub cmdHapus_Click()
Dim strHapusKode As String

strHapusKode = txtkdcust.Text
If MsgBox("Yakin ingin menghapus kode customer : " & strHapusKode & " ?", vbQuestion + vbOKCancel, "Hapus Customer") = 1 Then
    Con.Execute ("DELETE FROM customer WHERE kdcust='" & strHapusKode & "'")
    MsgBox "Data berhasil dihapus!", vbInformation, "Hapus"
    Call Reset
End If
End Sub

Private Sub cmdPrintKA_Click()
frmCetak.Show
'----cetak ke form
frmCetak.Font = "courier new"
frmCetak.CurrentX = 0
frmCetak.CurrentY = 0
frmCetak.FontSize = 9.5
frmCetak.Print Tab(3); "--------------------------------------";
frmCetak.Print Tab(3); "             KARTU ANGGOTA            ";
frmCetak.Print Tab(3); "--------------------------------------";
frmCetak.Print Tab(3); "No Anggota : "; frmcust.txtkdcust.Text;
frmCetak.Print Tab(3); "Nama       : "; frmcust.txtnmcust.Text;
frmCetak.Print Tab(3); "______________________________________";

'Dim prt As Printer
'
'For Each prt In Printers
'    If prt.DeviceName = "Generic / Text Only" Then
'        Set Printer = prt
'        Exit For
'    End If
'Next
'-----cetak ke printer
Printer.PaperSize = vbPRPSA4
Printer.Font = "courier new"
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.FontSize = 9.5
Printer.Print Tab(3); "--------------------------------------";
Printer.Print Tab(3); "             KARTU ANGGOTA            ";
Printer.Print Tab(3); "--------------------------------------";
Printer.Print Tab(3); "No Anggota : "; frmcust.txtkdcust.Text;
Printer.Print Tab(3); "Nama       : "; frmcust.txtnmcust.Text;
Printer.Print Tab(3); "______________________________________";
Printer.EndDoc
End Sub

Private Sub cmdSimpan_Click()
Dim sCust As String
Dim sTgl As String

sTgl = Format(dtTglMsk.Value, "yyyymmdd")
sCust = LTrim(txtkdcust.Text)
sCust = RTrim(sCust)

If txtkdcust.Text = Empty Then
    MsgBox "Kode customer masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtnmcust.Text = Empty Then
    MsgBox "Nama customer masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtalamat.Text = Empty Then
    MsgBox "Alamat masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf cbjenkel.Text = Empty Or cbjenkel.Text = "---------------Pilih---------------" Then
    MsgBox "Jenis kelamin masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txthp.Text = Empty Then
    MsgBox "Nomor HP masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
ElseIf dtTglMsk.Value = Empty Then
    MsgBox "Tanggal masuk masih kosong!, silakan isi dulu.", vbCritical, "Perhatian!"
    Exit Sub
End If

sSQL = "SELECT * FROM customer WHERE kdcust='" & sCust & "'"
RsCust.Open sSQL, Con, 3, 1

If RsCust.EOF Then
    sSQLInput = "INSERT INTO customer(kdcust, nmcust, alamat, jenkel, nohp, tglmsk) VALUES('" & sCust & "', '" & txtnmcust.Text & "', '" & txtalamat.Text & "', '" & cbjenkel.Text & "', '" & CStr(txthp.Text) & "', '" & sTgl & "')"
    'MsgBox (sSQLInput)
    MousePointer = 11
    Con.Execute (sSQLInput)
    MsgBox "Input data customer berhasil!", vbInformation, "Input"
    MousePointer = 0
    RsCust.Close
Else
    MsgBox "Kode customer sudah ada didatabase!", vbExclamation, "Perhatian!"
    RsCust.Close
    Exit Sub
End If

txtkdcust.Enabled = False
txtnmcust.Enabled = False
txtalamat.Enabled = False
cbjenkel.Enabled = False
txthp.Enabled = False
dtTglMsk.Enabled = False

txtkdcust.BackColor = &H80000013
txtnmcust.BackColor = &H80000013
txtalamat.BackColor = &H80000013
cbjenkel.BackColor = &H80000013
txthp.BackColor = &H80000013
'txttglmsk.BackColor = &H80000013

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
txtkdcust.SetFocus
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call KonekDB
Call Reset
End Sub

Private Sub txtalamat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    cbjenkel.SetFocus
End If
End Sub

Private Sub txtkdcust_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txtnmcust.SetFocus
End If
End Sub

Private Sub txtkdcust_LostFocus()
txtkdcust.Text = UCase(txtkdcust.Text)
End Sub


Private Sub txtnmcust_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txtalamat.SetFocus
End If
End Sub

Private Sub txthp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(";") Or KeyAscii = vbKeySpace Or KeyAscii = 13) Then
    txtalamat.Text = Empty
    MsgBox "Harap masukkan nomor HP dengan angka!", vbCritical, "Perhatian!"
    Beep
    KeyAscii = 0
    Exit Sub
End If
End Sub

Sub Reset()
txtkdcust.Text = Empty
txtnmcust.Text = Empty
txtalamat.Text = Empty
cbjenkel.Text = "---------------Pilih---------------"
txthp.Text = Empty
'dtTglMsk.Value = Empty

txtkdcust.Enabled = False
txtnmcust.Enabled = False
txtalamat.Enabled = False
cbjenkel.Enabled = False
txthp.Enabled = False
dtTglMsk.Enabled = False

txtkdcust.BackColor = &H80000013
txtnmcust.BackColor = &H80000013
txtalamat.BackColor = &H80000013
cbjenkel.BackColor = &H80000013
txthp.BackColor = &H80000013
'txttglmsk.BackColor = &H80000013

cmdBatal.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = False
cmdBersih.Enabled = False
cmdEdit.Enabled = False
End Sub

Sub UnReset()
txtkdcust.Enabled = True
txtnmcust.Enabled = True
txtalamat.Enabled = True
cbjenkel.Enabled = True
txthp.Enabled = True
dtTglMsk.Enabled = True

txtkdcust.BackColor = &H80000005
txtnmcust.BackColor = &H80000005
txtalamat.BackColor = &H80000005
cbjenkel.BackColor = &H80000005
txthp.BackColor = &H80000005
'txttglmsk.BackColor = &H80000005
End Sub



