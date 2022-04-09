VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParam 
   Caption         =   "Parameter Tanggal"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   Icon            =   "frmParam.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   2280
      MouseIcon       =   "frmParam.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdreport 
      Caption         =   "&Report"
      Height          =   495
      Left            =   600
      MouseIcon       =   "frmParam.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtTglakhir 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16121857
      CurrentDate     =   41622
   End
   Begin MSComCtl2.DTPicker dtTglmulai 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16121857
      CurrentDate     =   41622
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal Akhir"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal Mulai"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdreport_Click()
Dim ConRpt As New ADODB.Connection
Dim sMulai, sAkhir As String

sMulai = Format(dtTglmulai.Value, "yyyy-mm-dd")
sAkhir = Format(dtTglakhir.Value, "yyyy-mm-dd")

DataE_Pj.KoneksiPJ.ConnectionString = ConRpt.ConnectionString
DataE_Pj.Tbl_TRX sMulai, sAkhir
drTransaksi.Show
Unload Me
End Sub

Private Sub Form_Load()
dtTglmulai.Value = Now()
dtTglakhir.Value = Now()
End Sub
