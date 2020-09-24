VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmShitjaDitore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Report"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "frmShitjaDitore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   53
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   53
   End
   Begin MSMask.MaskEdBox dtprej 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox dtderi 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtdata 
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Sales Report"
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
      TabIndex        =   9
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type sales date"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmShitjaDitore.frx":058A
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Date:"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Prej:"
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
      TabIndex        =   6
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Deri:"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   495
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
Attribute VB_Name = "frmShitjaDitore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Public conn As New ADODB.Connection
Public rec As New ADODB.Recordset
Public cmd As New ADODB.Command
Private Sub Command1_Click()
On Error Resume Next
If dtprej = "" Then
flag = MsgBox("Type correct date?", vbYesNo + vbQuestion, "Warning!")
If flag = vbYes Then
If Rs.State = adStateOpen Then
Rs.Close
End If
Rs.Open "Select * from tblshitja", conn
Set rptShitja.DataSource = Rs
Load rptShitja
rptShitja.Show
Unload Me
Else
Exit Sub
dtprej.SetFocus
End If
Else
If Rs.State = adStateOpen Then
Rs.Close
End If
Rs.Open "select*from tblshitja where data between '" & dtprej.Text & "'and '" & dtderi.Text & "'", conn
Set rptShitja.DataSource = Rs
rptShitja.Show 1
Unload Me
End If
conn.Close
End Sub

Private Sub Form_Load()
txtdata.Text = Format(Now, "dd/mm/yyyy")
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\data\dbshitja.mdb"
conn.Open
cmd.ActiveConnection = conn
End Sub

Private Sub txtdata_Change()
dtprej.Text = txtdata.Text
dtderi.Text = txtdata.Text
End Sub
