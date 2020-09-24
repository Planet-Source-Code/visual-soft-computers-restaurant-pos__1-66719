VERSION 5.00
Begin VB.Form frmPerdoruesit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Users"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   Icon            =   "frmPerdoruesit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdRuaj 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
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
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
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
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
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
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdFshije 
      Caption         =   "Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   53
   End
   Begin VB.ComboBox txtniveli 
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
      ItemData        =   "frmPerdoruesit.frx":058A
      Left            =   1800
      List            =   "frmPerdoruesit.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
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
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
   End
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   53
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User Management"
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
      TabIndex        =   15
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Username and Password"
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmPerdoruesit.frx":05A7
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label3 
      Caption         =   "Level:"
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
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
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
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
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
Attribute VB_Name = "frmPerdoruesit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDalja_Click()
Unload Me
End Sub

Private Sub cmdKerko_Click()
FRMkERKOpERDORUESIT.Show 1

End Sub

Private Sub cmdNdrysho_Click()
txtemri.SetFocus
cmdRuaj.Caption = "Update"
cmdRuaj.Enabled = True
End Sub

Private Sub cmdRuaj_Click()
Select Case cmdRuaj.Caption
Case "Save"
Call dblidhja
With ar
.Open "Select *From tblPerdoruesit", strConek, adOpenStatic, adLockOptimistic
.AddNew
!emri = txtemri.Text
!fjalekalimi = txtpass.Text
!niveli = txtniveli.Text
.Update
MsgBox "Record saved successfully.", vbInformation, "Save!"
.Close
End With
Case "Update"
Call dblidhja
With ar
criteria = "Select *From tblPerdoruesit Where emri='" & txtemri & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
!emri = txtemri.Text
!fjalekalimi = txtpass.Text
!niveli = txtniveli.Text
.Update
MsgBox "Record updated successfully.", vbInformation, "Update!"
.Close
End With
End Select
End Sub

Private Sub cmdShto_Click()
txtemri.SetFocus
cmdRuaj.Enabled = True
End Sub

Private Sub txtemri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call dblidhja
With ar
criteria = "Select *From tblperdoruesit Where emri='" & txtemri & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
If .RecordCount = 1 Then
txtemri = !emri
txtpass = !fjalekalimi
txtniveli = !niveli
Else
MsgBox "Record not found.", vbInformation, "Not found!"
Exit Sub
End If
.Close
End With
End If
cmdFshije.Enabled = True
End Sub
