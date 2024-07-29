VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Data Barang"
   ClientHeight    =   9990
   ClientLeft      =   7290
   ClientTop       =   1905
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9990
   ScaleWidth      =   9315
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5040
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   3840
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5760
      Top             =   3720
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1335
      Left            =   480
      TabIndex        =   18
      Top             =   6360
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2355
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   8400
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IMPORT"
      Height          =   855
      Left            =   5280
      TabIndex        =   16
      Top             =   9000
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   480
      TabIndex        =   14
      Top             =   5280
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4260
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton chapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cubah 
      Caption         =   "Ubah"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton ctambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   480
   End
   Begin VB.TextBox tjumlah 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox tharga 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox tbarang 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3000
      Width           =   5055
   End
   Begin VB.TextBox tkode 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO TRANSAKSI"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATA BARANG UD. BAROKAH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label ljam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "jam"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label ltanggal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "tanggal"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jml Barang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Harga @"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chapus_Click()
Dim kode, barang, harga, jumlah As String

Open "c:\penjualan\databarang.dat" For Input As #1
Open "c:\penjualan\ganti.dat" For Append As #2

lewatkan:
    Do Until EOF(1)
        Input #1, kode, barang, harga, jumlah
        If UCase(kode) = UCase(tkode) Then
        GoTo lewatkan
        End If
        Write #2, ode, barang, harga, jumlah
    Loop
    Close
    
Kill "c:\penjualan\databarang.dat"
Name "c:\penjualan\ganti.dat" As "c:\penjualan\databarang.dat"
MsgBox "Data Sudah Diganti. Klik Ok!", vbOKOnly, "DATA BARANG UD. BAROKAH"
tkode = ""
tbarang = ""
tharga = ""
tjumlah = ""
tkode.SetFocus
Call tampildata

End Sub

Private Sub Command1_Click()
Form2.Visible = True


End Sub

Private Sub Command2_Click()
On Error GoTo segar
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim afile As String
    Dim i As Integer
    ListView2.ListItems.Clear

    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
    
    With CommonDialog1
        .DialogTitle = "Membuka Data"
        .CancelError = False
        .Filter = "HANYA FILE EXCEL" '(*.xls,*.xlsx)"
        .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    afile = .FileName
    End With
    
    ExcelObj.Workbooks.Open afile
    Text1.Text = afile
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)

    Dim l As ListItem
    
    With ExcelSheet
    i = 2
    Do Until .Cells(i, 1) & "" = ""
        Set l = ListView2.ListItems.Add(, , .Cells(i, 1))
        l.SubItems(1) = .Cells(i, 2)
        l.SubItems(2) = .Cells(i, 3)
        l.SubItems(3) = .Cells(i, 4)
        l.SubItems(4) = .Cells(i, 5)
        
               
        i = i + 1
    Loop

    End With

    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
    Timer2.Enabled = True
    ExcelObj.Workbooks.Close afile
   
    
segar:
'MsgBox "Ada kesalah dalam import data", vbInformation, "Data import guru"
Me.Refresh
End Sub

Private Sub ctambah_Click()
On Error GoTo biarkan
Open "C:\penjualan\databarang.dat" For Append As #1
Write #1, tkode, tbarang, tharga, tjumlah
Close #1
tkode = ""
tbarang = ""
tharga = ""
tjumlah = ""
tkode.SetFocus

Call tampildata
biarkan:
End Sub

Private Sub cubah_Click()
Dim kode, barang, harga, jumlah As String

Open "c:\penjualan\databarang.dat" For Input As #1
Open "c:\penjualan\ganti.dat" For Append As #2

Do Until EOF(1)
Input #1, kode, barang, harga, jumlah
If UCase(kode) = UCase(tkode) Then
kode = tkode
barang = tbarang
harga = tharga
jumlah = tjumlah
End If
Write #2, kode, barang, harga, jumlah
Loop
Close
Kill "c:\penjualan\databarang.dat"
Name "c:\penjualan\ganti.dat" As "c:\penjualan\databarang.dat"
MsgBox "Data Sudah Diganti, Klik OK", vbOKOnly, "DATA BARANG"
tkode = ""
tbarang = ""
tharga = ""
tjumlah = ""
tkode.SetFocus
Call tampildata


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

ListView2.View = lvwReport
ListView2.GridLines = True
ListView2.Sorted = True

ListView2.ColumnHeaders.Add , , "status", ListView1.Width / 12
ListView2.ColumnHeaders.Add , , "Kode", ListView1.Width / 12
ListView2.ColumnHeaders.Add , , "Nama Barang", ListView1.Width / 2
ListView2.ColumnHeaders.Add , , "Harga @", ListView1.Width / 5
ListView2.ColumnHeaders.Add , , "Jumlah", ListView1.Width / 4
Call tampildata
End Sub

Private Sub ListView1_DblClick()
tkode = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
tbarang = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
tharga = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(3)
tjumlah = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(4)

End Sub

Private Sub Timer1_Timer()
ljam.Caption = Format(Now, "HH:MM:SS")
ltanggal.Caption = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub tampildata()
Dim li As ListItem
Dim kode, nama, harga, jumlah As String
ListView1.ListItems.Clear
ListView1.GridLines = True
ListView1.Sorted = True

Open "c:\penjualan\databarang.dat" For Input As #1

Do Until EOF(1)
Input #1, kode, nama, harga, jumlah
Set li = ListView1.ListItems.Add(, , "ok")
li.SubItems(1) = kode
li.SubItems(2) = nama
li.SubItems(3) = harga
li.SubItems(4) = jumlah
Loop
Close #1

End Sub

Private Sub Timer2_Timer()
Dim o As Integer
Dim a As Integer
Dim kode, barang, harga, jumlah As String

On Error GoTo segarkan
For o = 1 To ListView2.ListItems.Count

 tkode = ListView2.ListItems(o).SubItems(1)
 tbarang = ListView2.ListItems(o).SubItems(2)
 tharga = ListView2.ListItems(o).SubItems(3)
 tjumlah = ListView2.ListItems(o).SubItems(4)
 
  a = 1
Open "C:\penjualan\databarang.dat" For Input As #1
Do Until EOF(1)
        Input #1, kode, barang, harga, jumlah
        If UCase(kode) = UCase(tkode) Then
        a = 0
        Text2.Text = a
        End If
Loop
Close
 If Text2.Text <> "0" Then
 Call ctambah_Click
 a = 1
 End If
 
 Next o
 
segarkan:
Call tampildata
Timer2.Enabled = False
tkode = ""
tbarang = ""
tharga = ""
tjumlah = ""
End Sub
