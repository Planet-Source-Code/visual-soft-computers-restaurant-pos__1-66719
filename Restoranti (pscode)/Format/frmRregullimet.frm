VERSION 5.00
Begin VB.Form frmRregullimet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "frmRregullimet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnulo 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   53
   End
   Begin VB.Frame Frame2 
      Caption         =   "Restaurant Informations"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   5055
      Begin VB.TextBox txtnumri 
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
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   4575
      End
      Begin VB.TextBox txtemri 
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
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Business number"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "Restaurant name"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Startup (Under Construction)"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
      Begin VB.CheckBox chStart 
         Caption         =   "Start program with Windows"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   4455
      End
   End
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   53
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmRregullimet.frx":038A
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Settings"
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
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   4335
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
Attribute VB_Name = "frmRregullimet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnulo_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If chStart.Value = 0 Then
WriteINI App.Path & "\Restauranti.ini", "StartWin", "StartWin", "0"
Else
WriteINI App.Path & "\Restauranti.ini", "StartWin", "StartWin", "1"
End If
'================
WriteINI App.Path & "\Restauranti.ini", "Emri", "Emri", txtemri.Text
WriteINI App.Path & "\Restauranti.ini", "Numri", "Numri", txtnumri.Text
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Arq1 As String
Dim Arq2 As String
Dim Arq3 As String
Dim Arq4 As String
Arq1 = ReadINI(App.Path & "\Restauranti.ini", "Fjalekalimi", "Fjalekalimi")
Arq2 = ReadINI(App.Path & "\Restauranti.ini", "StartWin", "StartWin")
Arq3 = ReadINI(App.Path & "\Restauranti.ini", "Emri", "Emri")
Arq4 = ReadINI(App.Path & "\Restauranti.ini", "Numri", "Numri")
chStart.Value = Arq2
txtemri.Text = Arq3
txtnumri.Text = Arq4
End Sub
