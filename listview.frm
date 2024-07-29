VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton hapus 
      Caption         =   "HAPUS"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   3720
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3201
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PENJUALAN"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tampildata()
Dim li As ListItem
Dim tkode, tbarang, tharga, tjumlah As String
ListView1.ListItems.Clear
ListView1.GridLines = True
ListView1.Sorted = True

Open "c:\penjualan\transaksi.dat" For Input As #1

Do Until EOF(1)
Set li = ListView1.ListItems.Add(, , "ok")
li.SubItems(1) = tkode
li.SubItems(2) = tbarang
li.SubItems(3) = tharga
li.SubItems(4) = tjumlah
Loop
Close #1

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
ListView1.ColumnHeaders.Add , , "Jumlah", ListView1.Width / 4
Call tampildata

End Sub

Private Sub hapus_Click()
Dim kode, barang, harga, jumlah As String

Open "c:\penjualan\transaksi.dat" For Input As #1
Open "c:\penjualan\ganti.dat" For Append As #2

lewatkan:
    Do Until EOF(1)
        
        If UCase(kode) = UCase(tkode) Then
        GoTo lewatkan
        End If
        Write #2, ode, barang, harga, jumlah
    Loop
    Close
    
Kill "c:\penjualan\transaksi.dat"
Name "c:\penjualan\ganti.dat" As "c:\penjualan\transaksi.dat"
MsgBox "Data Sudah Diganti. Klik Ok!", vbOKOnly, "DATA transaksi"
tkode = ""
tbarang = ""
tharga = ""
tjumlah = ""

Call tampildata

End Sub

Private Sub ListView1_DblClick()
tkode = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
tbarang = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
tharga = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(3)
tjumlah = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(4)
End Sub

