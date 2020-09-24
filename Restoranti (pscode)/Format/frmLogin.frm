VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Restauranti.isButton isButton1 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "frmLogin.frx":058A
      Style           =   6
      Caption         =   "Cancel"
      IconAlign       =   1
      iNonThemeStyle  =   0
      ShowFocus       =   -1  'True
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
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
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   53
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   53
   End
   Begin VB.TextBox txtuser 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin Restauranti.isButton isButton2 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Icon            =   "frmLogin.frx":0B24
      Style           =   6
      Caption         =   "OK"
      IconAlign       =   1
      iNonThemeStyle  =   0
      ShowFocus       =   -1  'True
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
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
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":10BE
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Username and Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UserName As String
Public DatabasePath As String
Public cn As ADODB.Connection
Dim WrongLogin As Integer
Dim Rs As ADODB.Recordset
Public Sub Open_cn()
Set cn = New ADODB.Connection
cn.CursorLocation = adUseClient
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.Properties("Data Source") = App.Path & "\Data\dbaza.mdb"
cn.Open
End Sub
Public Sub Close_cn()
cn.Close
Set cn = Nothing
End Sub


Private Sub isButton1_Click()
Unload Me
End Sub

Private Sub isButton2_Click()
On Error GoTo errhandler
If WrongLogin = 2 Then
Call MsgBox("You tried 3 times. Access denied!", vbCritical, "Error!")
End
End If
WrongLogin = WrongLogin + 1
If txtuser.Text = "" Or IsNull(txtuser.Text) = True Then
Call MsgBox("Please type Username.", vbInformation, "Username")
txtuser.SetFocus
Exit Sub
End If
If txtpass.Text = "" Or IsNull(txtpass.Text) = True Then
Call MsgBox("Please type password.", vbInformation, "Password")
txtpass.SetFocus
Exit Sub
End If
Open_cn
Set Rs = New ADODB.Recordset
Rs.Open ("Select * from tblPerdoruesit Where emri= '" & txtuser.Text & "'"), cn, adOpenStatic, adLockOptimistic, _
adCmdText
If txtpass.Text <> Rs.Fields("fjalekalimi") Then
Call MsgBox("Invalid password", vbInformation, "Error!")
txtpass.Text = ""
txtpass.SetFocus
Exit Sub
Else
UserName = txtuser.Text
Unload Me
frmMain.Show
frmMain.StatusBar1.Panels(3).Text = Rs.Fields("emri")
End If
Close_cn
Exit Sub
errhandler:
Call MsgBox("Invalid username", vbInformation, "Error!")
txtuser.Text = ""
txtpass.Text = ""
txtuser.SetFocus
Exit Sub

End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
isButton2_Click
End If

End Sub
