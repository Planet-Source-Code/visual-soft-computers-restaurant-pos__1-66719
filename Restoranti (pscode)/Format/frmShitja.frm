VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmShitja 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "frmShitja.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtid 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.PictureBox pickerkimi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5040
      Picture         =   "frmShitja.frx":240044
      ScaleHeight     =   2985
      ScaleWidth      =   5025
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtKerko 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1360
         TabIndex        =   25
         Top             =   70
         Width           =   3495
      End
      Begin ComctlLib.ListView lvkerkimi 
         Height          =   2295
         Left            =   160
         TabIndex        =   24
         Top             =   490
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "imgkerkimi"
         SmallIcons      =   "imgkerkimi"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1095
      End
      Begin ComctlLib.ImageList imgkerkimi 
         Left            =   360
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmShitja.frx":271B1C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   3600
      Picture         =   "frmShitja.frx":271DD6
      ScaleHeight     =   2775
      ScaleWidth      =   7815
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   7815
      Begin Restauranti.isButton isButton1 
         Height          =   495
         Left            =   2280
         TabIndex        =   17
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Icon            =   "frmShitja.frx":2BB936
         Style           =   1
         Caption         =   "Cancel"
         IconAlign       =   1
         iNonThemeStyle  =   0
         FontColor       =   8421504
         FontHighlightColor=   8421504
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   7455
         Begin VB.TextBox txtkusuri 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Digital Readout Thick Upright"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   705
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtpagesa 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Digital Readout Thick Upright"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   705
            Left            =   1560
            TabIndex        =   13
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "KUSURI:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   3960
            TabIndex        =   16
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "PAGUAR:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
      End
      Begin Restauranti.isButton isButton2 
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Icon            =   "frmShitja.frx":2BBED0
         Style           =   1
         Caption         =   "OK"
         IconAlign       =   1
         iNonThemeStyle  =   0
         FontColor       =   8421504
         FontHighlightColor=   8421504
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8055
         Y1              =   2760
         Y2              =   2760
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   7335
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12938
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList4"
      SmallIcons      =   "ImageList4"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Qty"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Price"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "TOTAL"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.PictureBox picsasia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   6360
      Picture         =   "frmShitja.frx":2BC8E2
      ScaleHeight     =   1545
      ScaleWidth      =   2385
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtsasia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Text            =   "1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ESC - To Cancel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Quantity"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6000
      TabIndex        =   8
      Top             =   840
      Width           =   9135
   End
   Begin VB.TextBox TXTTOTAL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Digital Readout Thick Upright"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   8760
      Width           =   2895
   End
   Begin ComctlLib.ListView lv 
      Height          =   8415
      Left            =   6000
      TabIndex        =   2
      Top             =   1320
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   14843
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Category"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by ID:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   870
      Width           =   3015
   End
   Begin VB.Label lblid 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1254"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   13440
      TabIndex        =   22
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill NR:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   21
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblkamarieri 
      Alignment       =   2  'Center
      Caption         =   "Qemajl Osmani"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Kamarieri:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblemri 
      BackStyle       =   0  'Transparent
      Caption         =   "Restauranti 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4335
   End
   Begin ComctlLib.ImageList ImageList4 
      Left            =   0
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShitja.frx":2C8FAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList3 
      Left            =   14760
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShitja.frx":2C9264
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   8760
      Width           =   2655
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   13920
      Top             =   10680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShitja.frx":2C957E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   14640
      Top             =   10680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmShitja.frx":2C9C90
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmShitja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub RuajShitjen()
'ruajtja e shitjes ne baze
'***********************************************************
For ilst = 1 To ListView1.ListItems.Count
lstpershkrimi = ListView1.ListItems(ilst).Text
lstsasia = ListView1.ListItems(ilst).SubItems(1)
lstqmimi = ListView1.ListItems(ilst).SubItems(2)
lsttotal = ListView1.ListItems(ilst).SubItems(3)
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
Call dbshitja
ac.Open strConek
With ar
criteria = "Select *From tblShitja"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
.AddNew
'!nr = lblid.Caption
!pershkrimi = lstpershkrimi
!sasia = lstsasia
!qmimi = lstqmimi
!total = lsttotal
!punetori = lblkamarieri
!nr = lblid
!Data = Format(Date, "dd/mm/yyyy")
.Update
.Close
End With
'***********************************************************
Next
ListView1.ListItems.Clear
TXTTOTAL.Text = ""
End Sub

Private Sub Combo1_Click()
If Combo1.Text = ">> All category" Then
lv.ListItems.Clear
Call dblidhja
With ar
criteria = "Select *From tblArtikujt"
.Open criteria, strConek, 3, 3
Do While Not .EOF
lv.ListItems.Add , , !pershkrimi, 1, 1
lv.ListItems(lv.ListItems.Count).SubItems(2) = !kategoria
lv.ListItems(lv.ListItems.Count).SubItems(1) = Format(!qmimi, "###,###,##0.00")
.MoveNext
Loop
.Close
End With
Else
lv.ListItems.Clear
Call dblidhja
ac.Open strConek
With ar
criteria = "Select * From tblArtikujt Where kategoria like'" & Combo1 & "%'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
If .RecordCount >= 1 Then
.MoveFirst
Do While Not .EOF
lv.ListItems.Add , , !pershkrimi, 1, 1
lv.ListItems(lv.ListItems.Count).SubItems(2) = !kategoria
lv.ListItems(lv.ListItems.Count).SubItems(1) = Format(!qmimi, "###,###,##0.00")
.MoveNext
Loop
End If
End With
End If
End Sub

Private Sub Command1_Click()
lv.View = lvwList
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()
lv.View = lvwIcon
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
Dim Arq As String
Arq = ReadINI(App.Path & "\Restauranti.ini", "Emri", "Emri")
lblemri.Caption = Arq
Call dblidhja
With ar
criteria = "Select *From tblArtikujt"
.Open criteria, strConek, 3, 3
Do While Not .EOF
lv.ListItems.Add , , !pershkrimi, 1, 1
lv.ListItems(lv.ListItems.Count).SubItems(2) = !kategoria
lv.ListItems(lv.ListItems.Count).SubItems(1) = Format(!qmimi, "###,###,##0.00")
.MoveNext
Loop
.Close
End With
'Vendosja e kategorive==============
Combo1.AddItem ">> All category"
Call dblidhja
ar.Open "Select *From tblKategoria", strConek, adOpenStatic, adLockOptimistic
ar.MoveFirst
Do While Not ar.EOF
Combo1.AddItem ar!kategoria
ar.MoveNext
Loop
ar.Close
'Vendosja e numrit te paragonit============
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
Call dbshitja
ac.Open strConek
With ar
criteria = "Select *From tblShitja"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
If ar.RecordCount = 0 Then
lblid.Caption = 1
Else
.MoveLast
lblid.Caption = Val(!nr) + 1
.Close
End If
End With
End Sub

Private Sub isButton1_Click()
Picture1.Visible = False
txtpagesa.Text = ""
txtkusuri.Text = ""
End Sub

Private Sub isButton2_Click()
RuajShitjen
txtpagesa.Text = ""
txtkusuri.Text = ""
Picture1.Visible = False
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyF1
If MsgBox("Are you sure youwant to cancel transaction?", vbYesNo + vbQuestion, "Confirm!") = vbYes Then
ListView1.ListItems.Clear
TXTTOTAL.Text = ""
End If
Case vbKeyF3
pickerkimi.Visible = True
txtKerko.SetFocus
Case vbKeyDelete
If ListView1.ListItems.Count = 0 Then
MsgBox "There are no items on sales list", vbOKOnly, "Void!"
Else
minusamount = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
TXTTOTAL.Text = Format(CCur(TXTTOTAL.Text) - minusamount, "###,###,##0.00")
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End If
Case vbKeyF6
If ListView1.ListItems.Count = 0 Then
MsgBox "There are no items on sales list", vbInformation, "Pay!"
Else
Picture1.Visible = True
txtpagesa.SetFocus
End If
Case vbKeyEscape
Unload Me
Case vbKeyF10
frmLogin.Show 1
Case vbKeyF11
Unload Me
Case vbKeyF12
picsasia.Visible = True
txtsasia.SetFocus
txtsasia.Text = ""
End Select
End Sub

Private Sub lk_Click()
lv.ListItems.Clear
Call dblidhja
With ar
criteria = "Select *From tblArtikujt Where kategoria='" & lk.SelectedItem.Text & "'"
.Open criteria, strConek, 3, 3
Do While Not .EOF
lv.ListItems.Add , , !pershkrimi, 1, 1
lv.ListItems(lv.ListItems.Count).SubItems(2) = !kategoria
lv.ListItems(lv.ListItems.Count).SubItems(1) = Format(!qmimi, "###,###,##0.00")
.MoveNext
Loop
.Close
End With

End Sub

Private Sub lv_DblClick()
ListView1.ListItems.Add , , lv.ListItems(lv.SelectedItem.Index).Text, , 1
ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = txtsasia
ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = lv.ListItems(lv.SelectedItem.Index).SubItems(1)
ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = Format(txtsasia * CCur(lv.ListItems(lv.SelectedItem.Index).SubItems(1)), "###,###,##0.00")
txtsasia.Text = "1"
inttotal = ListView1.ListItems(ListView1.ListItems.Count).SubItems(3)
TXTTOTAL.Text = Val(TXTTOTAL.Text) + Val(inttotal)
lv.SetFocus
End Sub


Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyF1
If MsgBox("Are you sure youwant to cancel transaction?", vbYesNo + vbQuestion, "Confirm!") = vbYes Then
ListView1.ListItems.Clear
TXTTOTAL.Text = ""
End If
Case vbKeyF3
pickerkimi.Visible = True
txtKerko.SetFocus
Case vbKeyDelete
If ListView1.ListItems.Count = 0 Then
MsgBox "There are no items on sales list", vbOKOnly, "Void!"
Else
minusamount = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
TXTTOTAL.Text = Format(CCur(TXTTOTAL.Text) - minusamount, "###,###,##0.00")
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End If
Case vbKeyF6
If ListView1.ListItems.Count = 0 Then
MsgBox "There are no items on sales list", vbInformation, "Pay!"
Else
Picture1.Visible = True
txtpagesa.SetFocus
End If
Case vbKeyEscape
Unload Me
Case vbKeyF10
frmLogin.Show 1
Case vbKeyF11
Unload Me
Case vbKeyF12
picsasia.Visible = True
txtsasia.SetFocus
txtsasia.Text = ""
End Select
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ListView1.ListItems.Add , , lv.ListItems(lv.SelectedItem.Index).Text, , 1
ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = txtsasia
ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = lv.ListItems(lv.SelectedItem.Index).SubItems(1)
ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = Format(txtsasia * CCur(lv.ListItems(lv.SelectedItem.Index).SubItems(1)), "###,###,##0.00")
txtsasia.Text = "1"
inttotal = ListView1.ListItems(ListView1.ListItems.Count).SubItems(3)
TXTTOTAL.Text = Val(TXTTOTAL.Text) + Val(inttotal)
lv.SetFocus
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtid_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyF1
If MsgBox("Are you sure youwant to cancel transaction?", vbYesNo + vbQuestion, "Confirm!") = vbYes Then
ListView1.ListItems.Clear
TXTTOTAL.Text = ""
End If
Case vbKeyF3
pickerkimi.Visible = True
txtKerko.SetFocus
Case vbKeyDelete
If ListView1.ListItems.Count = 0 Then
MsgBox "There are no items on sales list", vbOKOnly, "Void!"
Else
minusamount = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
TXTTOTAL.Text = Format(CCur(TXTTOTAL.Text) - minusamount, "###,###,##0.00")
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End If
Case vbKeyF6
If ListView1.ListItems.Count = 0 Then
MsgBox "There are no items on sales list", vbInformation, "Pay!"
Else
Picture1.Visible = True
txtpagesa.SetFocus
End If
Case vbKeyEscape
Unload Me
Case vbKeyF10
frmLogin.Show 1
Case vbKeyF11
Unload Me
Case vbKeyF12
picsasia.Visible = True
txtsasia.SetFocus
txtsasia.Text = ""
End Select
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
Call dblidhja
ac.Open strConek
With ar
criteria = "Select *From tblartikujt Where id='" & txtid & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
ListView1.ListItems.Add , , !pershkrimi, , 1
ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = txtsasia
ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = Format(CCur(!qmimi), "###,###,##0.00")
ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = Format(txtsasia * CCur(!qmimi), "###,###,##0.00")
txtsasia.Text = "1"
inttotal = ListView1.ListItems(ListView1.ListItems.Count).SubItems(3)
TXTTOTAL.Text = Val(TXTTOTAL.Text) + Val(inttotal)
End With
txtid.Text = ""
End If
End Sub

Private Sub txtKerko_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
pickerkimi.Visible = False
End Select
End Sub

Private Sub txtKerko_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim inti As ListItem
If KeyAscii = 13 Then
Set ac = New ADODB.Connection
Set ar = New ADODB.Recordset
Call dblidhja
ac.Open strConek
lvkerkimi.ListItems.Clear
With ar
criteria = "Select *From tblArtikujt"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
.MoveFirst
Do While Not .EOF
If Mid(!pershkrimi, 1, Len(txtKerkimi)) = txtKerkimi Then
Set intitem = lvkerkimi.ListItems.Add(, , !pershkrimi, , 1)
intitem.SubItems(1) = !id
End If
.MoveNext
lvkerkimi.SetFocus
Loop
.Close
End With
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtpagesa_Change()
'txtkusuri.Text = Val(txtpagesa) - Val(TXTTOTAL)
If TXTTOTAL.Text = "" Then
Exit Sub
End If
If txtpagesa.Text = "" Then
txtkusuri.Text = ""
Else
txtkusuri.Text = Format(txtpagesa - CCur(TXTTOTAL), "###,###,##0.00")
End If
End Sub

Private Sub txtpagesa_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
Picture1.Visible = False
txtpagesa.Text = ""
txtkusuri.Text = ""
End Select

End Sub

Private Sub txtpagesa_KeyPress(KeyAscii As Integer)
'If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then KeyAscii = 0
If KeyAscii = 13 Then
isButton2_Click
End If
End Sub

Private Sub txtsasia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
picsasia.Visible = False
ListView1.SetFocus
End If
If KeyAscii = vbKeyEscape Then
picsasia.Visible = False
txtsasia.Text = "1"
txtid.SetFocus
End If
End Sub
