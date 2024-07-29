VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "TRANSAKSI"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7245
   LinkTopic       =   "Form2"
   ScaleHeight     =   8565
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ccancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton ctambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   4320
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cprint 
      Caption         =   "&PRINT"
      Height          =   615
      Left            =   360
      TabIndex        =   21
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox ttotal 
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox tbayar 
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox tkembali 
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   7800
      Width           =   1575
   End
   Begin VB.TextBox tjumlahbayar 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox tkode 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox tbarang 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2880
      Width           =   5055
   End
   Begin VB.TextBox tharga 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox tjumlah 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3840
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO DATA BARANG"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   360
      TabIndex        =   14
      Top             =   4920
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3201
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pembayaran"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jml Bayar"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSAKSI PENJUALAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Harga @"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jml Barang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label ltanggal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "tanggal"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label ljam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "jam"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Visible = True
Form2.Visible = False

End Sub

Private Sub cprint_Click()
Dim grs As String
On Error GoTo err_han
Me.CommonDialog1.CancelError = True
Me.CommonDialog1.ShowPrinter
grs = String$(40, "=")
Printer.Print
Printer.Print
Printer.FontName = "Times New Roman"
Printer.FontSize = 8
Printer.Print
Printer.Print
Printer.Print Tab(0); "STRUK PEMBAYARAN TUNAI"
Printer.Print Tab(0); "UD BAROKAH"
Printer.Print Tab(0); "UD BAROKAH"
Printer.Print Tab(0); "Jl. Raya Kalanganyar Barat 53 Sedati"
Printer.Print
Printer.FontName = "Arial Narrow"
Printer.FontSize = 7
Printer.Print Tab(0); ljam.Caption; Tab(16); ltanggal.Caption
Printer.Print Tab(0); grs
Printer.FontName = "arial narrow"
Printer.FontSize = 7
Printer.FontBold = False
For i = 1 To ListView1.ListItems.Count
    Printer.Print Tab(0); ListView1.ListItems(i).ListSubItems(1); _
    Tab(10); ListView1.ListItems(i).ListSubItems(2); _
    Tab(30); ListView1.ListItems(i).ListSubItems(3); _
    Tab(40); ListView1.ListItems(i).ListSubItems(4); _
    Tab(45); ListView1.ListItems(i).ListSubItems(5)
Next
Printer.Print Tab(0); grs
Printer.Print Tab(0); "Total Bayar"; Tab(14); ":"; Tab(16); ttotal.Text
Printer.Print Tab(0); "Bayar"; Tab(14); ":"; Tab(16); tbayar.Text
Printer.Print Tab(0); "Kembali"; Tab(14); ":"; Tab(16); tkembali.Text

Printer.Print Tab(30); "Petugas"
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(30); "_______________"

Printer.EndDoc
Exit Sub
err_han:
If Err.Number = 32755 Then
MsgBox "cetak dibatalkan", vbInformation
End If



    



End Sub

Private Sub ctambah_Click()
tkode.SetFocus
tkode.Text = ""
tharga.Text = ""
tbarang.Text = ""
tjumlah.Text = ""
End Sub

Private Sub Form_Load()
Dim li As ListItem
ListView1.View = lvwReport
ListView1.GridLines = True
ListView1.Sorted = True

ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 12
ListView1.ColumnHeaders.Add , , "Kode", ListView1.Width / 12
ListView1.ColumnHeaders.Add , , "Nama Barang", ListView1.Width / 2
ListView1.ColumnHeaders.Add , , "Harga @", ListView1.Width / 5
ListView1.ColumnHeaders.Add , , "Jumlah", ListView1.Width / 12
ListView1.ColumnHeaders.Add , , "Total Biaya", ListView1.Width / 4
End Sub

Private Sub tbayar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
tkembali = Val(tbayar) - Val(ttotal)
cprint.SetFocus
End If
If KeyCode = 27 Then End

End Sub

Private Sub Timer1_Timer()
ljam = Format(Now, "HH:MM:SS")
ltanggal = Format(Date, "DD/MM/YYYY")

If (tbarang <> "0" And tjumlah <> "0") Or (tbarang <> "" And tjumlah <> "") Then
tjumlahbayar.Text = Val(tharga) * Val(tjumlah)
Else
tjumlahbayar.Text = ""
End If

End Sub

Private Sub caridata()
Dim kode, barang, harga, jumlah As String
Dim ada As Integer

'On Error GoTo ulangi
Open "c:\penjualan\databarang.dat" For Input As #1

Do Until EOF(1)
Input #1, kode, barang, harga, jumlah
If UCase(tkode) = UCase(kode) Then
tkode = kode
tbarang = barang
tharga = harga
tjumlah = 1
tjumlah.SetFocus
ada = ada + 1
End If
Loop
Close
If ada <> 0 Then tkode.Enabled = True
If ada = 0 Then
MsgBox "Data Barang Tidak Ada", vbOKOnly, "Aplikasi Penjualan UD.Barokah"
'ulangi:
tkode.SetFocus
End If

End Sub


Private Sub tjumlah_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo biarkan
If KeyCode = 13 Then
Call tampilkandata
ctambah.SetFocus
End If
If KeyCode = 27 Then End
Call jumlahkan
biarkan:
End Sub

Private Sub tkode_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo biarkan
If KeyCode = 13 Then
Call caridata
End If
If KeyCode = 27 Then End
biarkan:
End Sub

Private Sub tampilkandata()
Dim li As ListItem

ListView1.View = lvwReport
ListView1.GridLines = True
ListView1.Sorted = True

Set li = ListView1.ListItems.Add(, , "ok")
li.SubItems(1) = tkode
li.SubItems(2) = tbarang
li.SubItems(3) = tharga
li.SubItems(4) = tjumlah
li.SubItems(5) = tjumlahbayar
tkode.SetFocus
End Sub

Private Sub jumlahkan()
On Error GoTo biarkan
ttotal.Text = ""
biarkan:
For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems(i).ListSubItems(5).Text <> "" Then ttotal.Text = Val(ttotal.Text) + ListView1.ListItems(i).ListSubItems(5).Text
Next i
End Sub

