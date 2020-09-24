VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmParagonet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Bills"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmParagonet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1400
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   6240
      Width           =   1695
   End
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lv 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   6615
      _ExtentX        =   11668
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
         Text            =   "Total"
         Object.Width           =   1411
      EndProperty
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   53
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   4080
      Top             =   1320
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
            Picture         =   "frmParagonet.frx":15162
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Bill NR:"
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmParagonet.frx":154B4
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Bills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Display sales bills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmParagonet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lv.ListItems.Clear
Call dbshitja
ac.Open strConek
With ar
criteria = "Select * From tblShitja Where nr like'" & Text1 & "%'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
Do While Not .EOF
lv.ListItems.Add , , !pershkrimi, 1, 1
lv.ListItems(lv.ListItems.Count).SubItems(1) = Format(!qmimi, "###,###,##0.00")
lv.ListItems(lv.ListItems.Count).SubItems(2) = sasia
lv.ListItems(lv.ListItems.Count).SubItems(3) = Format(!total, "###,###,##0.00")
.MoveNext
Loop
.Close
End With
End If
End Sub

