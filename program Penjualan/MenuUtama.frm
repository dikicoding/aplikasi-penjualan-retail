VERSION 5.00
Begin VB.MDIForm MenuUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Penjualan"
   ClientHeight    =   5100
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7485
   Icon            =   "MenuUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_master 
      Caption         =   "&Master"
      Begin VB.Menu mnuBrg 
         Caption         =   "&Barang"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuCustomer 
         Caption         =   "&Customer"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuUser 
         Caption         =   "&User"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuTrx 
      Caption         =   "&Transaksi"
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "&Report"
      Begin VB.Menu rptBarang 
         Caption         =   "&Barang"
      End
      Begin VB.Menu rptCustomer 
         Caption         =   "&Customer"
      End
      Begin VB.Menu rptUser 
         Caption         =   "&User"
      End
      Begin VB.Menu rptTransaksi 
         Caption         =   "&Transaksi"
      End
   End
   Begin VB.Menu mnuKeluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "MenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuBrg_Click()
frmbrg.Show
End Sub

Private Sub mnuCustomer_Click()
frmcust.Show
End Sub

Private Sub mnuKeluar_Click()
Unload Me
End Sub

Private Sub mnuTrx_Click()
frmTransaksi.Show
End Sub

Private Sub mnuUser_Click()
frmuser.Show
End Sub

Private Sub rptBarang_Click()
drBarang.Show
End Sub

Private Sub rptCustomer_Click()
drCustomer.Show
End Sub

Private Sub rptTransaksi_Click()
frmParam.Show
End Sub

Private Sub rptUser_Click()
drUser.Show
End Sub
