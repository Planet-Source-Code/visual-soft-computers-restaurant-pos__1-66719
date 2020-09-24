VERSION 5.00
Begin VB.Form frmFurnizuesit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier's"
   ClientHeight    =   5160
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmKlientet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtid 
      DataField       =   "emri"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
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
      Height          =   470
      Left            =   3720
      TabIndex        =   15
      Top             =   3960
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
      Height          =   470
      Left            =   1920
      TabIndex        =   14
      Top             =   3960
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
      Height          =   470
      Left            =   120
      TabIndex        =   13
      Top             =   3960
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
      Height          =   470
      Left            =   3720
      TabIndex        =   12
      Top             =   4560
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
      Height          =   470
      Left            =   1920
      TabIndex        =   11
      Top             =   4560
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
      Height          =   470
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   1695
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   53
   End
   Begin VB.TextBox txtemail 
      DataField       =   "email"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox txttelefoni 
      DataField       =   "telefoni"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox txtadresa 
      DataField       =   "adresa"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox txtemri 
      DataField       =   "emri"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1800
      Width           =   3495
   End
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   17
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   53
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers Record"
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
      TabIndex        =   19
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type supplier information and then click Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmKlientet.frx":058A
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID:"
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
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "E-Mail:"
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
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Phone Nr:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Address:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Full name:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
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
Attribute VB_Name = "frmFurnizuesit"
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
criteria = "Select *From tblFurnizuesit Where id='" & txtid & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
!id = txtid.Text
!emri = txtemri.Text
!adresa = txtadresa.Text
!telefoni = txttelefoni.Text
!email = txtemail.Text
!datelindja = txtdatelindja.Text
!vendlindja = txtvendlindja.Text
.Delete
MsgBox "Records deleted successfully.", vbInformation, "Delete!"
.Close
End With
End Sub

Private Sub cmdKerko_Click()
frmKerkoFur.Show 1
End Sub

Private Sub cmdNdrysho_Click()
cmdRuaj.Caption = "Update"
cmdRuaj.Enabled = True
End Sub

Private Sub cmdRuaj_Click()
Select Case cmdRuaj.Caption
Case "Save"
Call dblidhja
With ar
.Open "Select *From tblFurnizuesit", strConek, adOpenStatic, adLockOptimistic
.AddNew
!id = txtid.Text
!emri = txtemri.Text
!adresa = txtadresa.Text
!telefoni = txttelefoni.Text
!email = txtemail.Text
.Update
MsgBox "Records saved successfully.", vbInformation, "Save!"
.Close
End With
Case "Update"
Call dblidhja
With ar
criteria = "Select *From tblFurnizuesit Where id='" & txtid & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
!id = txtid.Text
!emri = txtemri.Text
!adresa = txtadresa.Text
!telefoni = txttelefoni.Text
!email = txtemail.Text
.Update
MsgBox "Record updated successfully.", vbInformation, "Save!"
.Close
End With
End Select
End Sub

Private Sub cmdShto_Click()
txtid.SetFocus
cmdRuaj.Enabled = True
End Sub

Private Sub txtemri_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call dblidhja
With ar
criteria = "Select *From tblFurnizuesit Where id='" & txtid & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
If .RecordCount = 1 Then
txtid = !id
txtemri = !emri
txtadresa = !adresa
txttelefoni = !telefoni
txtemail = !email
txtdatelindja = !datelindja
txtvendlindja = !vendlindja
Else
MsgBox "This record doesn't exist.", vbInformation, "Not found!"
Exit Sub
End If
.Close
End With
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
cmdFshije.Enabled = True
End Sub
