VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "TRANSAKSI"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15375
   LinkTopic       =   "Form2"
   ScaleHeight     =   10665
   ScaleWidth      =   15375
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   9480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture4 
      Height          =   375
      Left            =   6360
      Picture         =   "kasir benr.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   46
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Left            =   5760
      Picture         =   "kasir benr.frx":1DF2
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   45
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   5160
      Picture         =   "kasir benr.frx":3BE4
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   44
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox ttotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1095
      IMEMode         =   3  'DISABLE
      Left            =   12720
      TabIndex        =   16
      Text            =   "0"
      Top             =   960
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   5160
      ScaleHeight     =   11.25
      ScaleLeft       =   50
      ScaleMode       =   0  'User
      ScaleTop        =   20
      ScaleWidth      =   9.75
      TabIndex        =   41
      Top             =   845
      Width           =   255
   End
   Begin VB.CommandButton simpan 
      Caption         =   "SIMPAN"
      Height          =   735
      Left            =   11280
      TabIndex        =   24
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cari 
      Caption         =   "CARI DATA"
      Height          =   855
      Left            =   11280
      TabIndex        =   23
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "kasir benr.frx":59D6
      Left            =   8760
      List            =   "kasir benr.frx":59E3
      TabIndex        =   22
      Text            =   "Pilih"
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   10560
      Top             =   10200
   End
   Begin VB.CommandButton ccancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   11280
      TabIndex        =   21
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton ctambah 
      Caption         =   "Tambah"
      Height          =   735
      Left            =   11280
      TabIndex        =   20
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cprint 
      BackColor       =   &H00000000&
      Caption         =   "&PRINT"
      Height          =   735
      Left            =   12480
      MaskColor       =   &H00404040&
      TabIndex        =   19
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox tbayar 
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox tkembali 
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   7800
      Width           =   1575
   End
   Begin VB.TextBox tjumlahbayar 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   7800
      Width           =   1575
   End
   Begin VB.TextBox tkode 
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox tbarang 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   5760
      Width           =   5055
   End
   Begin VB.TextBox tharga 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox tjumlah 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   7320
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10080
      Top             =   10200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO DATA BARANG"
      Height          =   855
      Left            =   12480
      TabIndex        =   0
      Top             =   8520
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5741
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
   Begin VB.Label ljam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "jam"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   13200
      TabIndex        =   5
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Reguler"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   13200
      TabIndex        =   59
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Ònline"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   11280
      TabIndex        =   58
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "RINGKASAN PROMO"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12000
      TabIndex        =   51
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11280
      TabIndex        =   50
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   11280
      TabIndex        =   49
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label27 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   11160
      TabIndex        =   48
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Height          =   1695
      Left            =   11160
      TabIndex        =   47
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "MEMBER"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5640
      TabIndex        =   43
      Top             =   825
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rp"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   11160
      TabIndex        =   42
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "301-2804VT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   40
      Top             =   1500
      Width           =   3375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "WAHID R"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   39
      Top             =   900
      Width           =   3375
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7320
      TabIndex        =   38
      Top             =   1440
      Width           =   3540
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF00&
      Caption         =   "       User"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   37
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label16 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   5040
      TabIndex        =   36
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "301-2804VT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   35
      Top             =   1499
      Width           =   3375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "WAHID R"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   34
      Top             =   900
      Width           =   3375
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFF00&
      Caption         =   " Faktur"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   33
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFF00&
      Caption         =   "       User"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   32
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label17 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   11160
      TabIndex        =   31
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label15 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   30
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "K42 (SC17)"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   13200
      TabIndex        =   29
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "V.2017.4.1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "--Dev--"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   7080
      TabIndex        =   27
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "POS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6000
      TabIndex        =   26
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   15375
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   7800
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
      TabIndex        =   12
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   5760
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
      Top             =   5760
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
      Top             =   6240
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
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label ltanggal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "tanggal"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   11400
      TabIndex        =   6
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   57
      Top             =   7200
      Width           =   10695
   End
   Begin VB.Label Label35 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   56
      Top             =   7080
      Width           =   10935
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   11280
      TabIndex        =   60
      Top             =   10200
      Width           =   3735
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   55
      Top             =   5640
      Width           =   6375
   End
   Begin VB.Label Label33 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   54
      Top             =   5520
      Width           =   6615
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   7080
      TabIndex        =   53
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Label Label31 
      BackColor       =   &H000000FF&
      Height          =   1455
      Left            =   6960
      TabIndex        =   52
      Top             =   5520
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub simpantransaksi()
Open "c:\penjualan\transaksi.dat" For Append As #1
Write #1, ltanggal, ljam, tkode, tbarang, tharga, tjumlah, tjumlahbayar
Close #1
tkode = ""
tbarang = ""
tharga = ""
tjumlah = ""
tstok = ""
tkode.SetFocus
'Call tampildata
Exit Sub

Beep
End Sub
Private Sub caridata2()
Dim kode, barang, harga, jumlah As String

'On Error GoTo ulangi
Open "c:\penjualan\databarang.dat" For Input As #1

Do Until EOF(1)
    
    If UCase(tkode) = UCase(kode) Then
       tstok = jumlah
    End If
Loop
Close
End Sub
Private Sub rubah()
Dim kode, barang, harga, jumlah As String

Open "c:\penjualan\databarang.dat" For Input As #1
Open "c:\penjualan\ganti.dat" For Append As #2


    Do Until EOF(1)
        
        If UCase(kode) = UCase(tkode) Then
        kode = tkode
        barang = tbarang
        harga = tharga
        tstok.Text = Val(tstok) - Val(tjumlah)
        jumlah = tstok
        
        
        End If
        Write #2, kode, barang, harga, jumlah
    Loop
    Close
Kill "c:\penjualan\databarang.dat"
Name "c:\penjualan\ganti.dat" As "c:\penjualan\databarang.dat"
End Sub



Private Sub Combo1_Click()
  If Combo1.Text = "1" Then
    tkode.Caption = "1"
    
    Else
    Combo1.Text = "2"
    tkode.Caption = "2"
    End If
End Sub

Private Sub Command1_Click()
Form1.Visible = True
Form2.Visible = True


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
Timer2.Enabled = True
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

Private Sub simpan_Click()
On Error GoTo biarkan
Open "C:\penjualan\transaksi.dat" For Append As #1
Write #1, tkode, tbarang, tharga, tjumlah
Close #1
tkode = ""
tbarang = ""
tharga = ""
tjumlah = ""
tkode.SetFocus

'Call tampildata
biarkan:
End Sub

Private Sub tbayar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
tkembali = Val(tbayar) - Val(ttotal)
'cprint.SetFocus'
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


Private Sub Timer2_Timer()
Dim o As Integer
Dim kode, barang, harga, jumlah As String

On Error GoTo segarkan
For o = 1 To ListView1.ListItems.Count

 tkode = ListView1.ListItems(o).SubItems(1)
 tbarang = ListView1.ListItems(o).SubItems(2)
 tharga = ListView1.ListItems(o).SubItems(3)
 tjumlah = ListView1.ListItems(o).SubItems(4)
  Call caridata2
  Call rubah
  Call simpantransaksi
Next o
 Timer2.Enabled = False
 
segarkan:
'Call tampildata
'Timer2.Enabled = False
End Sub

Private Sub tjumlah_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo biar
If KeyCode = 13 Then
Call tampilkandata
ctambah.SetFocus
End If
If KeyCode = 27 Then End
Call jumlahkan
biar:
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

