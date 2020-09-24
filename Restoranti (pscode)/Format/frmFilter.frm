VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter records by Supplier"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   6120
      Width           =   1815
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   53
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
   End
   Begin ComctlLib.ListView lv 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList2"
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Price"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Qty"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Supplier"
         Object.Width           =   3528
      EndProperty
   End
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   53
   End
   Begin VB.Label Label3 
      Caption         =   "Supplier:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Supplier from Combo list"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter by Supplier"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmFilter.frx":058A
      Top             =   120
      Width           =   810
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   7560
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFilter.frx":2A52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Call dblidhja
ac.Open strConek
With ar
criteria = "Select * From tblArtikujt Where furnizuesi like'" & Combo1 & "%'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
Do While Not .EOF
lv.ListItems.Add , , !pershkrimi, 1, 1
lv.ListItems(lv.ListItems.Count).SubItems(1) = Format(!qmimi, "###,###,##0.00")
lv.ListItems(lv.ListItems.Count).SubItems(2) = !sasia
lv.ListItems(lv.ListItems.Count).SubItems(3) = !furnizuesi
.MoveNext
Loop
.Close
End With
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call dblidhja
ar.Open "Select *From tblArtikujt", strConek, adOpenStatic, adLockOptimistic
ar.MoveFirst
Do While Not ar.EOF
Combo1.AddItem ar!furnizuesi
ar.MoveNext
Loop
ar.Close
End Sub
