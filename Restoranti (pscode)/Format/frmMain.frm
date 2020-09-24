VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Restaurant Point of Sales 1.0"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10665
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   8280
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   2858
      ButtonWidth     =   2408
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "SALE"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   27
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Products"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Suppliers"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Category"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   26
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Products Delivery"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sales Bills"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   28
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Users"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   720
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7905
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "www.visualsoftdev.com"
            TextSave        =   "www.visualsoftdev.com"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "Username:"
            TextSave        =   "Username:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Qemajl Osmani"
            TextSave        =   "Qemajl Osmani"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   10605
      TabIndex        =   1
      Top             =   1620
      Width           =   10665
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   28
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20850E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":209160
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":209DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20AA04
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20B656
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20C2A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20CEFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20DB4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20E79E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20F3F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":210042
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":210C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2118E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":212538
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21318A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":213DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":214A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":215680
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2162D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":216F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":217B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2187C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21941A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21A06C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21ACBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21B910
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21C562
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21D1B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSkedar 
      Caption         =   "File"
      Begin VB.Menu mnuShit 
         Caption         =   "SALE"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuP 
         Caption         =   "Select printer"
      End
      Begin VB.Menu mnuDalja 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuArt 
      Caption         =   "Products"
      Begin VB.Menu mnuShto 
         Caption         =   "Add new product"
      End
      Begin VB.Menu mnuFilteri 
         Caption         =   "Filter by supplier"
      End
   End
   Begin VB.Menu mnuKat 
      Caption         =   "Category"
      Begin VB.Menu mnuKatArt 
         Caption         =   "Category"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Other records"
      Begin VB.Menu mnuFur 
         Caption         =   "Supplier"
      End
      Begin VB.Menu mnuFurnizimi 
         Caption         =   "Delivery"
      End
      Begin VB.Menu mnuPar 
         Caption         =   "Sales bill's"
      End
   End
   Begin VB.Menu mnuRaportet 
      Caption         =   "Reports"
      Begin VB.Menu mnuDit 
         Caption         =   "Daily Report"
      End
      Begin VB.Menu mnuPeriodik 
         Caption         =   "Periodic Report"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiss 
         Caption         =   "Product List"
      End
   End
   Begin VB.Menu mnuPerd 
      Caption         =   "Users"
      Begin VB.Menu mnuPerdoruesit 
         Caption         =   "Users"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "About"
      Begin VB.Menu mnuRreth 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PictureButton10_Click()
frmShitja.Show 1
Unload Me
End Sub

Private Sub mnuDalja_Click()
Unload Me
End Sub

Private Sub mnuDit_Click()
frmShitjaDitore.Show 1
End Sub

Private Sub mnuFilteri_Click()
frmFilter.Show 1
End Sub

Private Sub mnuFur_Click()
frmFurnizuesit.Show 1
End Sub

Private Sub mnuFurnizimi_Click()
frmFurnizimi.Show 1
End Sub

Private Sub mnuKatArt_Click()
frmKategoria.Show 1
End Sub

Private Sub mnuLiss_Click()
Set dB = New Connection
dB.CursorLocation = adUseClient
dB.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\dbaza.mdb;" & ";Persist Security Info=False;Jet OLEDB:Database Password=cc03bn01"
Set adoLista = New Recordset
adoLista.Open "select * from tblArtikujt order by id ASC;", dB, adOpenStatic, adLockOptimistic
Set rptLista.DataSource = adoLista
rptLista.Show 1
Unload Me
End Sub

Private Sub mnuLista_Click()

End Sub

Private Sub mnuP_Click()
cd1.ShowPrinter
End Sub

Private Sub mnuPar_Click()
frmParagonet.Show 1
End Sub

Private Sub mnuPerdoruesit_Click()
frmPerdoruesit.Show 1
End Sub

Private Sub mnuPeriodik_Click()
frmReportShitja.Show 1
End Sub

Private Sub mnuRreth_Click()
frmInfo.Show 1
End Sub

Private Sub mnuShit_Click()
frmShitja.Show
End Sub

Private Sub mnuShto_Click()
frmArtikujt.Show 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1:
frmShitja.lblkamarieri.Caption = frmMain.StatusBar1.Panels(3).Text
frmShitja.Show 1
Case 2:
frmArtikujt.Show 1
Case 3:
frmFurnizuesit.Show 1
Case 4:
frmKategoria.Show 1
Case 5:
frmFurnizimi.Show 1
Case 6:
frmParagonet.Show
Case 7: PopupMenu mnuRaportet, , Button.Left, (Button.Top + Button.Height)
Case 8:
frmRregullimet.Show 1
Case 9:
frmPerdoruesit.Show 1
Case 10:
frmInfo.Show 1
Case 11:
Unload Me
End Select
End Sub
