VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   53
   End
   Begin VB.PictureBox ctrlLiner2 
      BackColor       =   &H00C0C0C0&
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   1440
      Width           =   6015
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
      Left            =   2880
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Computers   www.digitalcomputers.net"
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
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "www.visualsoftdev.com                    info@visualsoftdev.com"
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
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright: Visual Soft Development 2000-2006"
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Restaurant Point of Sales 1.0"
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   240
      Picture         =   "frmInfo.frx":058A
      Top             =   0
      Width           =   4020
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1470
      Left            =   -360
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
