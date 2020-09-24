VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmKategoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Category"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   Icon            =   "frmKategoria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   53
   End
   Begin Restauranti.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   53
   End
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   2880
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7080
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Baza.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Baza.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblKategoria"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox ctrlLiner2 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   7335
      TabIndex        =   7
      Top             =   1200
      Width           =   7335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   1680
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton cmdDalja 
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
      Height          =   400
      Left            =   3480
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdKerko 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdFshije 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdRuaje 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdNdrysho 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3480
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdShto 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
   Begin ComctlLib.ImageList imgkerkimi 
      Left            =   3000
      Top             =   5760
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
            Picture         =   "frmKategoria.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Category name:"
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
      TabIndex        =   12
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category records"
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
      TabIndex        =   9
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      TabIndex        =   8
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmKategoria.frx":0844
      Top             =   120
      Width           =   810
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   6480
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
            Picture         =   "frmKategoria.frx":2D0C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmKategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDalja_Click()
Unload Me
End Sub

Private Sub cmdFshije_Click()
Call dblidhja
With ar
criteria = "Select *From tblkategoria Where kategoria='" & Text1 & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
!kategoria = Text1.Text
.Delete
MsgBox "Category deleted successfully.", vbInformation, "Delete!"
.Close
End With
End Sub

Private Sub cmdKerko_Click()
frmKerkoKat.Show 1
End Sub

Private Sub cmdNdrysho_Click()
cmdRuaje.Caption = "Update"
cmdRuaje.Enabled = True
End Sub

Private Sub cmdRuaje_Click()
Select Case cmdRuaje.Caption
Case "Save"
Call dblidhja
With ar
.Open "Select *From tblkategoria", strConek, adOpenStatic, adLockOptimistic
.AddNew
!kategoria = Text1.Text
.Update
MsgBox "Category saved successfully.", vbInformation, "Save!"
Text1.Text = ""
.Close
End With
Case "Update"
Call dblidhja
With ar
criteria = "Select *From tblkategoria Where kategoria='" & Text1 & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
!kategoria = Text1.Text
.Update
MsgBox "Category updated successfully.", vbInformation, "Update!"
Text1.Text = ""
.Close
End With
End Select
End Sub

Private Sub cmdShto_Click()
Text1.SetFocus
End Sub

