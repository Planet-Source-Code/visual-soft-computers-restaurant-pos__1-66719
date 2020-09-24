VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmArtikujt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Management"
   ClientHeight    =   6555
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmArtikujt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtfurnizuesi 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Width           =   3375
   End
   Begin VB.ComboBox txtkategoria 
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
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   3375
   End
   Begin Restauranti.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   23
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   53
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
      TabIndex        =   22
      Top             =   6000
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
      TabIndex        =   21
      Top             =   6000
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
      TabIndex        =   20
      Top             =   6000
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
      TabIndex        =   19
      Top             =   5400
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
      TabIndex        =   18
      Top             =   5400
      Width           =   1695
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
      TabIndex        =   17
      Top             =   5400
      Width           =   1695
   End
   Begin Restauranti.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   53
   End
   Begin VB.TextBox txtafati 
      DataField       =   "tvsh"
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtvat 
      DataField       =   "tvsh"
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtqmimi 
      DataField       =   "qmimi"
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtsasia 
      DataField       =   "sasia"
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtpershkrimi 
      DataField       =   "pershkrimi"
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtid 
      DataField       =   "id"
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type item information and then click Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   25
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Management"
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
      TabIndex        =   24
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmArtikujt.frx":038A
      Top             =   120
      Width           =   810
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   5400
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
            Picture         =   "frmArtikujt.frx":2852
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sales Price:"
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
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblLabels 
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
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "VAT:"
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
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Price:"
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
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Quantity:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Category:"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Description:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
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
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
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
Attribute VB_Name = "frmArtikujt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cleartext()
txtid.Text = ""
txtpershkrimi.Text = ""
txtfurnizuesi.Text = ""
txtkategoria.Text = ""
txtqmimi.Text = ""
txtsasia.Text = ""
txtvat.Text = ""
txtafati.Text = ""
End Sub
Private Sub cmdDalja_Click()
Unload Me
End Sub

Private Sub cmdFshije_Click()
Call dblidhja
With ar
criteria = "Select *From tblartikujt Where id='" & txtid & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
!id = txtid.Text
!pershkrimi = txtpershkrimi.Text
!kategoria = txtkategoria.Text
!sasia = txtsasia.Text
!qmimi = txtqmimi.Text
!tvsh = txtvat.Text
!furnizuesi = txtfurnizuesi.Text
!qmimi_shitjes = txtafati.Text
.Delete
MsgBox "Record saved successfully.", vbInformation, "Save!"
.Close
End With
End Sub

Private Sub cmdKerko_Click()
frmKerkoArtikujt.Show 1
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
.Open "Select *From tblArtikujt", strConek, adOpenStatic, adLockOptimistic
.AddNew
!id = txtid.Text
!pershkrimi = txtpershkrimi.Text
!kategoria = txtkategoria.Text
!sasia = txtsasia.Text
!qmimi = txtqmimi.Text
!tvsh = txtvat.Text
!furnizuesi = txtfurnizuesi.Text
!qmimi_shitjes = txtafati.Text
.Update
MsgBox "Record saved successfully.", vbInformation, "Update!"
.Close
End With
Case "Update"
Call dblidhja
With ar
criteria = "Select *From tblartikujt Where id='" & txtid & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
!id = txtid.Text
!pershkrimi = txtpershkrimi.Text
!kategoria = txtkategoria.Text
!sasia = txtsasia.Text
!qmimi = txtqmimi.Text
!tvsh = txtvat.Text
!furnizuesi = txtfurnizuesi.Text
!qmimi_shitjes = txtafati.Text
.Update
MsgBox "Record updated successfully.", vbInformation, "Update!"
.Close
End With
End Select

End Sub

Private Sub cmdShto_Click()
Call cleartext
txtid.SetFocus
cmdRuaj.Enabled = True
End Sub

Private Sub Form_Load()
'Add Category in combobox
Call dblidhja
ar.Open "Select *From tblKategoria", strConek, adOpenStatic, adLockOptimistic
ar.MoveFirst
Do While Not ar.EOF
txtkategoria.AddItem ar!kategoria
ar.MoveNext
Loop
ar.Close
'Add supplier in combobox
Call dblidhja
ar.Open "Select *From tblFurnizuesit", strConek, adOpenStatic, adLockOptimistic
ar.MoveFirst
Do While Not ar.EOF
txtfurnizuesi.AddItem ar!emri
ar.MoveNext
Loop
ar.Close
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call dblidhja
With ar
criteria = "Select *From tblArtikujt Where id='" & txtid & "'"
.Open criteria, strConek, adOpenStatic, adLockOptimistic
If .RecordCount = 1 Then
txtid = !id
txtpershkrimi = !pershkrimi
txtkategoria = !kategoria
txtsasia = !sasia
txtqmimi = !qmimi
txtvat = !tvsh
txtfurnizuesi = !furnizuesi
txtafati = !qmimi_shitjes
Else
MsgBox "Record not found.", vbInformation, "Not found!"
Exit Sub
End If
.Close
End With
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
cmdFshije.Enabled = True
End Sub

Private Sub txtpershkrimi_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtvat_LostFocus()
txtafati.Text = Format(txtqmimi / 100 * txtvat + txtqmimi, "###,###,##0.00")
End Sub
