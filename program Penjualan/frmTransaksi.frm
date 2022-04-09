VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransaksi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Penjualan"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   Icon            =   "frmTransaksi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   12285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPesan 
      Caption         =   "Pesan"
      Height          =   855
      Left            =   9600
      TabIndex        =   29
      Top             =   1800
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvPesanan 
      Height          =   2655
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   4683
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Barang"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Harga Barang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Diskon"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total Harga"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Header Transaksi"
      Height          =   2175
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   5655
      Begin VB.TextBox txtkdtrx 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtkdcust 
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtnmcust 
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   22
         Top             =   1680
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker dtTglTrx 
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   93454337
         CurrentDate     =   41617
      End
      Begin VB.Label Label11 
         Caption         =   "Kode Transaksi"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Customer"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Customer"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Tgl Transaksi"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "bln/tgl/tahun"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Input Pesanan"
      Height          =   2175
      Left            =   5880
      TabIndex        =   9
      Top             =   720
      Width           =   6375
      Begin VB.TextBox txtdiskon 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txthrgbrg 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtnmbrg 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtkdbrg 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtqty 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Diskon"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Harga Barang"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Nama Barang"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Kode Barang"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Qty"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   5415
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         Height          =   495
         Left            =   2760
         MouseIcon       =   "frmTransaksi.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Bata&l"
         Height          =   495
         Left            =   1440
         MouseIcon       =   "frmTransaksi.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "T&ambah"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmTransaksi.frx":0A56
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   2760
         MouseIcon       =   "frmTransaksi.frx":0D60
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   4080
         MouseIcon       =   "frmTransaksi.frx":106A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBersih 
         Caption         =   "&Bersih"
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmTransaksi.frx":1374
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
         Height          =   495
         Left            =   1440
         MouseIcon       =   "frmTransaksi.frx":167E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label lblGrandTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   33
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label12 
      Caption         =   "Grand Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   32
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Transaksi Penjualan"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmTransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTrx As New ADODB.Recordset
Dim sSQL As String
Dim sSQLInput As String
Dim mylist As ListItem
Dim iFlag As Integer
Dim dSumGT As Double

Private Sub cmdbatal_Click()
Call Reset
cmdCari.Enabled = True
cmdTambah.Enabled = True
lblGrandTotal.Caption = Empty
End Sub

Private Sub cmdBersih_Click()
txtkdtrx.Text = Empty
txtkdcust.Text = Empty
txtnmcust.Text = Empty
txtkdbrg.Text = Empty
txtnmbrg.Text = Empty
txthrgbrg.Text = Empty
txtqty.Text = Empty
txtdiskon.Text = Empty
txtkdtrx.Enabled = True
txtkdcust.Enabled = True
cmdCari.Enabled = False
cmdTambah.Enabled = False
lvPesanan.ListItems.Clear
lblGrandTotal.Caption = Empty
'txtkdcust.SetFocus
End Sub

Private Sub cmdCari_Click()
Dim strCari As String
Dim sSQLC As String
Dim sSQLDetail As String
Dim RsCari As New ADODB.Recordset
Dim RsDetail As New ADODB.Recordset
Dim RsCust As New ADODB.Recordset
Dim sSQLCust As String
Dim dGrandTotal As Double

strCari = InputBox("Masukkan kode transaksi :", "Cari transaksi")

sSQLC = "SELECT * FROM transaksi WHERE kdtrx='" & strCari & "'"
RsCari.Open sSQLC, Con, 3, 1

If strCari = Empty Then
    MsgBox "Kode transaksi masih kosong!", vbCritical, "Perhatian"
ElseIf RsCari.EOF Then
    MsgBox "Data tidak diketemukan!", vbExclamation, "Cari"
Else
    txtkdtrx.Text = RsCari!kdtrx
    dtTglTrx.Value = RsCari!tgltrx
    txtkdcust.Text = RsCari!kdcust
    
    sSQLCust = "SELECT * FROM customer WHERE kdcust='" & txtkdcust.Text & "'"
    RsCust.Open sSQLCust, Con, 3, 1

    txtnmcust.Text = RsCust!nmcust
    RsCust.Close
    sSQLDetail = "SELECT * FROM detailtrx a INNER JOIN barang b ON a.kdbrg=b.kdbrg WHERE kdtrx='" & strCari & "'"
    RsDetail.Open sSQLDetail, Con, 3, 1
    
    lvPesanan.ListItems.Clear
    
    Do Until RsDetail.EOF
        Set mylist = lvPesanan.ListItems.Add(, , RsDetail!kdbrg)
        mylist.SubItems(1) = RsDetail!nmbrg
        mylist.SubItems(2) = RsDetail!hrgbrg
        mylist.SubItems(3) = RsDetail!qty
        mylist.SubItems(4) = RsDetail!disc
        mylist.SubItems(5) = (RsDetail!qty * RsDetail!hrgbrg) - ((RsDetail!qty * RsDetail!hrgbrg) * (RsDetail!disc) / 100)
        dGrandTotal = dGrandTotal + mylist.SubItems(5)
        RsDetail.MoveNext
    Loop
    
    lblGrandTotal.Caption = Format(dGrandTotal, "Rp #,#.00")
    Call UnReset
    txtkdcust.Enabled = False
    cmdHapus.Enabled = True
    cmdBatal.Enabled = True
    txtkdtrx.Enabled = False
    dtTglTrx.Enabled = False
    iFlag = 1
    RsDetail.Close
End If
RsCari.Close
Call EnDis_Pesanan(False)
End Sub

Sub EnDis_Pesanan(X As Boolean)
txtkdbrg.Enabled = X
txtnmbrg.Enabled = X
txthrgbrg.Enabled = X
txtqty.Enabled = X
txtdiskon.Enabled = X
End Sub

Private Sub cmdHapus_Click()
Dim strHapusKode As String

strHapusKode = txtkdtrx.Text
If MsgBox("Yakin ingin menghapus kode transaksi: " & strHapusKode & " ?", vbQuestion + vbOKCancel, "Hapus Transaksi") = 1 Then
    Con.Execute ("DELETE FROM transaksi WHERE kdtrx='" & strHapusKode & "'")
    Con.Execute ("DELETE FROM detailtrx WHERE kdtrx='" & strHapusKode & "'")
    MsgBox "Data berhasil dihapus!", vbInformation, "Hapus"
    Call Reset
End If
End Sub

Private Sub cmdPesan_Click()
Dim RsCekStok As New ADODB.Recordset
Dim sSQLCekStok As String

If txtkdbrg.Text = Empty Then
    MsgBox "Kode barang masih kosong!, silakan isi dulu", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtqty.Text = Empty Then
    MsgBox "Qty masih kosong!, silakan isi dulu", vbCritical, "Perhatian!"
    Exit Sub
ElseIf txtdiskon.Text = Empty Then
    MsgBox "Diskon masih kosong!, silakan isi dulu", vbCritical, "Perhatian!"
Else
    sSQLCekStok = "SELECT satuan, stok FROM barang WHERE kdbrg='" & txtkdbrg.Text & "'"
    RsCekStok.Open sSQLCekStok, Con, 3, 1
    If RsCekStok!stok = 0 Then
        MsgBox "Stok Kosong!", vbExclamation, "Cek Stok"
        RsCekStok.Close
        Exit Sub
    ElseIf CInt(RsCekStok!stok) < CInt(txtqty.Text) Then
        MsgBox "Stok tidak mencukupi!, stok yang tersedia hanya " & RsCekStok!stok & " " & RsCekStok!satuan, vbExclamation, "Cek Stok"
        RsCekStok.Close
        Exit Sub
    End If
    
    Set mylist = lvPesanan.ListItems.Add(, , txtkdbrg.Text)
    mylist.SubItems(1) = txtnmbrg.Text
    mylist.SubItems(2) = txthrgbrg.Text
    mylist.SubItems(3) = txtqty.Text
    mylist.SubItems(4) = txtdiskon.Text
    mylist.SubItems(5) = (txthrgbrg.Text * txtqty.Text) - ((txthrgbrg.Text * txtqty.Text) * (txtdiskon.Text) / 100)
        
    dSumGT = dSumGT + CDbl(mylist.SubItems(5))
    lblGrandTotal.Caption = Format(dSumGT, "Rp #,#.00")
    txtkdbrg.SelStart = 0
    txtkdbrg.SelLength = Len(txtkdbrg.Text)
    txtkdbrg.SetFocus
    RsCekStok.Close
End If
End Sub

Private Sub cmdSimpan_Click()
Dim iCekVal As Integer
Dim RsDetail As New ADODB.Recordset
Dim RsPotStok As New ADODB.Recordset
Dim RsSelectStok As New ADODB.Recordset
Dim sSQLSelect As String
Dim sSQLPotStok As String
Dim iPtgStok As Integer

iCekVal = lvPesanan.ListItems.Count

sSQL = "SELECT * FROM transaksi WHERE kdtrx='" & txtkdtrx.Text & "'"
RsTrx.Open sSQL, Con, 3, 1

If iCekVal = 0 Then
    MsgBox "Detail pesanan masih kosong!", vbExclamation, "Informasi"
    RsTrx.Close
    Exit Sub
ElseIf RsTrx.EOF Then
    sSQLInput = "INSERT INTO transaksi(kdtrx, tgltrx, kdcust) VALUES('" & txtkdtrx.Text & "', '" & Format(dtTglTrx.Value, "yyyy-mm-dd") & "', '" & txtkdcust.Text & "')"
    MousePointer = 11
    Con.Execute (sSQLInput)
    
    RsDetail.Open "detailtrx", Con, adOpenKeyset, adLockOptimistic
    i = 1
    Do While i <> iCekVal + 1
        RsDetail.AddNew
        RsDetail!kdtrx = txtkdtrx.Text
        RsDetail!kdbrg = lvPesanan.ListItems(i)
        sSQLSelect = "SELECT kdbrg, stok FROM barang WHERE kdbrg='" & lvPesanan.ListItems(i) & "'"
        RsSelectStok.Open sSQLSelect, Con, 3, 1
        'MsgBox RsSelectStok!kdbrg & " " & RsSelectStok!stok
        RsDetail!qty = lvPesanan.ListItems(i).SubItems(3)
        sSQLPotStok = "UPDATE barang SET stok='" & CInt(RsSelectStok!stok) - CInt(lvPesanan.ListItems(i).SubItems(3)) & "' WHERE kdbrg='" & lvPesanan.ListItems(i) & "'"
        RsPotStok.Open sSQLPotStok, Con, 3, 1
        RsDetail!disc = lvPesanan.ListItems(i).SubItems(4)
        RsDetail.Update
        RsSelectStok.Close
        i = i + 1
    Loop
'    RsPotStok.Close
    RsDetail.Close
    MsgBox "Input data transaksi berhasil!", vbInformation, "Input Transaksi"
    MousePointer = 0
    RsTrx.Close
'    RsPotStok.Close
Else
    MsgBox "Kode transaksi sudah ada!", vbExclamation, "Perhatian!"
    RsTrx.Close
    Exit Sub
End If

dSumGT = 0
txtkdtrx.Enabled = False
txtkdcust.Enabled = False
txtnmcust.Enabled = False
txtkdbrg.Enabled = False
txtkdbrg.Enabled = False
txtnmbrg.Enabled = False
txthrgbrg.Enabled = False
txtqty.Enabled = False
txtdiskon.Enabled = False
dtTglTrx.Enabled = False

txtkdtrx.BackColor = &H80000013
txtkdcust.BackColor = &H80000013
txtnmcust.BackColor = &H80000013
txtkdbrg.BackColor = &H80000013
txtkdbrg.BackColor = &H80000013
txtnmbrg.BackColor = &H80000013
txthrgbrg.BackColor = &H80000013
txtqty.BackColor = &H80000013
txtdiskon.BackColor = &H80000013
'txttglmsk.BackColor = &H80000013

cmdBatal.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = True
cmdBersih.Enabled = False
cmdCari.Enabled = True
cmdTambah.Enabled = True
End Sub

Private Sub cmdTambah_Click()
iFlag = 0
cmdBatal.Enabled = True
cmdSimpan.Enabled = True
cmdBersih.Enabled = True
cmdCari.Enabled = False
cmdHapus.Enabled = False
cmdPesan.Enabled = True
Call UnReset
Call cmdBersih_Click
cmdTambah.Enabled = False
txtkdtrx.SetFocus
End Sub

Private Sub cmdTutup_Click()
Unload Me
End Sub

Private Sub dtTglTrx_Change()
txtkdcust.SetFocus
End Sub

Private Sub Form_Load()
dtTglTrx.Value = Now()
Call KonekDB
Call Reset
iFlag = 0
dSumGT = 0
End Sub

Private Sub txtdiskon_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(";") Or KeyAscii = vbKeySpace Or KeyAscii = 13) Then
    MsgBox "Harap masukkan diskon dengan angka!", vbCritical, "Perhatian!"
    Beep
    KeyAscii = 0
    Exit Sub
ElseIf KeyAscii = 13 Then
    cmdPesan.SetFocus
End If
End Sub

Private Sub txtkdbrg_KeyPress(KeyAscii As Integer)
Dim RsBrg As New ADODB.Recordset
Dim sSQLBrg As String

sSQLBrg = "SELECT nmbrg, hrgbrg FROM barang WHERE kdbrg='" & txtkdbrg.Text & "'"
RsBrg.Open sSQLBrg, Con, 3, 1

KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    If txtkdbrg.Text = Empty Then
        MsgBox "Kode barang masih kosong!", vbCritical, "Perhatian"
    ElseIf RsBrg.EOF Then
        MsgBox "Data tidak diketemukan!", vbExclamation, "Informasi"
    Else
        txtnmbrg.Text = RsBrg!nmbrg
        txthrgbrg.Text = RsBrg!hrgbrg
        txtqty.SetFocus
    End If
End If
RsBrg.Close
End Sub

Private Sub txtkdcust_KeyPress(KeyAscii As Integer)
Dim RsCust As New ADODB.Recordset
Dim sSQLCust As String

sSQLCust = "SELECT * FROM customer WHERE kdcust='" & txtkdcust.Text & "'"
RsCust.Open sSQLCust, Con, 3, 1

KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    If txtkdcust.Text = Empty Then
        MsgBox "Kode customer masih kosong!", vbCritical, "Perhatian"
    ElseIf RsCust.EOF Then
        MsgBox "Data tidak diketemukan!", vbExclamation, "Informasi"
    Else
        txtnmcust.Text = RsCust!nmcust
        txtkdcust.Enabled = False
        If iFlag = 0 Then
            txtkdbrg.SetFocus
        End If
    End If
End If
RsCust.Close
End Sub

Private Sub txtkdcust_LostFocus()
txtkdcust.Text = UCase(txtkdcust.Text)
End Sub

Private Sub txtkdtrx_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txtkdtrx.Enabled = False
    dtTglTrx.SetFocus
End If
End Sub

Private Sub txtkdtrx_LostFocus()
txtkdtrx.Text = UCase(txtkdtrx.Text)
End Sub

Private Sub txtnmcust_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 Then
    txtkdbrg.SetFocus
End If
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(";") Or KeyAscii = vbKeySpace Or KeyAscii = 13) Then
    'txtkdbrg.Text = Empty
    MsgBox "Harap masukkan nomor HP dengan angka!", vbCritical, "Perhatian!"
    Beep
    KeyAscii = 0
    Exit Sub
ElseIf KeyAscii = 13 Then
    txtdiskon.SetFocus
End If
End Sub

Sub Reset()
txtkdtrx.Text = Empty
txtkdcust.Text = Empty
txtnmcust.Text = Empty
txtkdbrg.Text = Empty
txtnmbrg.Text = Empty
txthrgbrg.Text = Empty
txtqty.Text = Empty
txtdiskon.Text = Empty
'dtTglTrx.Value = Empty

txtkdtrx.Enabled = False
txtkdcust.Enabled = False
txtnmcust.Enabled = False
txtkdbrg.Enabled = False
txtnmbrg.Enabled = False
txthrgbrg.Enabled = False
txtqty.Enabled = False
txtdiskon.Enabled = False
dtTglTrx.Enabled = False

txtkdtrx.BackColor = &H80000013
txtkdcust.BackColor = &H80000013
txtnmcust.BackColor = &H80000013
txtkdbrg.BackColor = &H80000013
txtnmbrg.BackColor = &H80000013
txthrgbrg.BackColor = &H80000013
txtqty.BackColor = &H80000013
txtdiskon.BackColor = &H80000013
'txttglmsk.BackColor = &H80000013

cmdBatal.Enabled = False
cmdSimpan.Enabled = False
cmdHapus.Enabled = False
cmdBersih.Enabled = False
cmdPesan.Enabled = False

lvPesanan.ListItems.Clear
End Sub

Sub UnReset()
txtkdtrx.Enabled = True
txtkdcust.Enabled = True
txtnmcust.Enabled = False
txtkdbrg.Enabled = True
txtnmbrg.Enabled = False
txthrgbrg.Enabled = False
txtqty.Enabled = True
txtdiskon.Enabled = True
dtTglTrx.Enabled = True

txtkdtrx.BackColor = &H80000005
txtkdcust.BackColor = &H80000005
txtnmcust.BackColor = &H80000005
txtkdbrg.BackColor = &H80000005
txtnmbrg.BackColor = &H80000005
txthrgbrg.BackColor = &H80000005
txtqty.BackColor = &H80000005
txtdiskon.BackColor = &H80000005
'txttglmsk.BackColor = &H80000005
End Sub



